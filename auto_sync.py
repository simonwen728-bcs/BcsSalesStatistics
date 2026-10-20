"""Sync Customer/Pull on the first use of each America/New_York day.
The completion stamp and a short-lived claim live in the business database.
No cron thread, local stamp file, credentials, or paid Render worker is needed.
"""
import hashlib
import json
import os
import threading
import time
import uuid
from datetime import datetime
from zoneinfo import ZoneInfo

KEY = 'masterDailySync'
ZONE = ZoneInfo('America/New_York')
MANAGER = None
MANAGER_LOCK = threading.Lock()


def today():
    return datetime.now(ZONE).date().isoformat()


def enabled():
    return os.environ.get('AUTO_SYNC', '0').strip() == '1'


class DailySync:
    def __init__(self, server):
        self.server = server
        self.lock = threading.RLock()
        self.thread = None
        self.closed = False
        self.completed_day = ''
        self.retry_after = 0
        self.source = hashlib.sha256(server.QUOTE_PULL_API_URL.encode()).hexdigest()

    def _read(self, con):
        suffix = ' FOR UPDATE' if self.server.DATABASE_URL else ''
        row = con.execute('SELECT value FROM bcs_prefs WHERE key=?' + suffix, (KEY,)).fetchone()
        value = json.loads(row[0]) if row else {}
        # A changed source should get its own successful daily synchronization.
        if value.get('source') not in (None, self.source):
            return {}
        return value

    def _store(self, con, value):
        con.execute('INSERT INTO bcs_prefs(key,value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value',
                    (KEY, json.dumps(value, ensure_ascii=False)))

    def maybe_start(self):
        if not enabled():
            return False
        day = today()
        with self.lock:
            if self.closed or self.completed_day == day or time.time() < self.retry_after or (self.thread and self.thread.is_alive()):
                return False
            now = time.time()
            owner = uuid.uuid4().hex
            with self.server.DB_LOCK:
                con = self.server.connect()
                try:
                    if not self.server.DATABASE_URL:
                        con.execute('BEGIN IMMEDIATE')
                    con.execute('INSERT INTO bcs_prefs(key,value) VALUES (?,?) ON CONFLICT(key) DO NOTHING', (KEY, '{}'))
                    value = self._read(con)
                    if value.get('successDay') == day:
                        self.completed_day = day
                        con.commit()
                        return False
                    if value.get('leaseUntil', 0) > now or value.get('retryAfter', 0) > now:
                        con.commit()
                        return False
                    value.update(source=self.source, owner=owner, leaseUntil=now + 600,
                                 attemptDay=day, lastAttemptAt=now, running=True, lastError='')
                    self._store(con, value)
                    con.commit()
                finally:
                    con.close()
            self.thread = threading.Thread(target=self._perform, args=(day, owner), daemon=True, name='bcs-daily-master-sync')
            self.thread.start()
            return True

    def _perform(self, day, owner):
        result = None
        try:
            result = self.server.API.sync_master_from_sheets('all')
            if not isinstance(result, dict):
                result = {'ok': False, 'error': '同步返回格式错误'}
        except Exception as error:
            result = {'ok': False, 'error': str(error)}
        now = time.time()
        try:
            with self.server.DB_LOCK:
                con = self.server.connect()
                try:
                    value = self._read(con)
                    if value.get('owner') != owner:
                        return
                    value.update(running=False, leaseUntil=0)
                    if result.get('ok'):
                        value.update(successDay=day, lastSuccessAt=now, retryAfter=0, lastError='',
                                     customers=result.get('customers', 0), stock=result.get('stock', 0),
                                     changed=bool(result.get('changed', True)))
                    else:
                        value.update(lastError=str(result.get('error') or '同步失败')[:500], retryAfter=now + 60)
                    self._store(con, value)
                    con.commit()
                finally:
                    con.close()
            with self.lock:
                if result.get('ok'):
                    self.completed_day = day
                else:
                    self.retry_after = now + 60
            message = '客户=%s 库存=%s' % (result.get('customers', 0), result.get('stock', 0)) if result.get('ok') else str(result.get('error', '同步失败'))
            print('[daily-master-sync] %s %s %s' % (day, 'success' if result.get('ok') else 'failed', message), flush=True)
        except Exception as error:
            # The lease expires on its own if persisting the completion stamp fails.
            print('[daily-master-sync] 状态保存失败: %s' % error, flush=True)

    def status(self):
        with self.server.DB_LOCK:
            con = self.server.connect()
            try:
                row = con.execute('SELECT value FROM bcs_prefs WHERE key=?', (KEY,)).fetchone()
                value = json.loads(row[0]) if row else {}
                revision = self.server.read_data_revision(con)
            finally:
                con.close()
        public = {k: value[k] for k in ('successDay', 'lastAttemptAt', 'lastSuccessAt', 'lastError', 'customers', 'stock', 'changed') if k in value}
        return dict(public, enabled=enabled(), timezone='America/New_York', today=today(),
                    running=bool(value.get('running') and value.get('leaseUntil', 0) > time.time()), revision=revision)

    def stop(self):
        with self.lock:
            self.closed = True
            thread = self.thread
        if thread and thread.is_alive():
            thread.join(timeout=1)


def start_worker(server):
    global MANAGER
    with MANAGER_LOCK:
        if MANAGER is None or MANAGER.server is not server or MANAGER.closed:
            MANAGER = DailySync(server)
        return MANAGER


def maybe_start(server):
    try:
        return start_worker(server).maybe_start()
    except Exception as error:
        # A temporary sync/state error must not prevent querying the saved data.
        print('[daily-master-sync] 自动同步暂未启动: %s' % error, flush=True)
        return False


def get_status(server):
    return start_worker(server).status()
