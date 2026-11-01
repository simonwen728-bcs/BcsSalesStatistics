"""Inspect first; repair missing entries/names using the same role and an unambiguous source."""
from collections import defaultdict
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re


def clean(value):
    return str(value or '').strip()


def norm(value):
    return re.sub(r'[^a-z0-9]', '', clean(value).lower())


def valid_name(value, code):
    text = clean(value)
    return bool(text and text.upper() != clean(code).upper()
                and norm(text) not in {'repname','rep1name','rep2name','repcode','customerid','customername','exsalesrep','itemledgerno'}
                and not re.fullmatch(r'C\d+', text, re.I))


def snapshot(con):
    sales = [tuple(r) for r in con.execute('SELECT id,rep1,rep1_name,rep2,rep2_name FROM sales ORDER BY id')]
    reps = [tuple(r) for r in con.execute('SELECT kind,code,name,group_,locked,deleted FROM reps ORDER BY kind,code')]
    return {'sales': sales, 'reps': reps}


def plan(state):
    from sales_import_validation import valid_rep
    directory = {kind: {} for kind in ('s1','s2')}
    names = defaultdict(set)
    for kind, code, name, group, locked, deleted in state['reps']:
        if kind not in directory:
            continue
        code = clean(code).upper()
        directory[kind][code] = (name, group, locked, deleted)
        if not deleted and valid_rep(code, name) and valid_name(name, code):
            names[kind,code].add(clean(name))
    seen = {kind: set() for kind in directory}
    blank_codes = {'s1':0,'s2':0}
    invalid = set()
    for rid, c1, n1, c2, n2 in state['sales']:
        for kind, code, name in (('s1',c1,n1),('s2',c2,n2)):
            code = clean(code).upper()
            if not code:
                blank_codes[kind] += int(kind != "s2" or clean(c1).upper() != "IC")  # BCS_IC_REP2_EXEMPT
                continue
            if not valid_rep(code, name):
                invalid.add((kind,code))
                continue
            seen[kind].add(code)
            if valid_name(name, code):
                names[kind,code].add(clean(name))
    new = []
    for kind in ('s1','s2'):
        for code in sorted(seen[kind] - set(directory[kind])):
            choices = names[kind,code]
            new.append({'kind':kind,'code':code,'name':next(iter(choices)) if len(choices)==1 else ''})
    fixes = []
    unresolved = 0
    for rid, c1, n1, c2, n2 in state['sales']:
        for kind, code, name in (('s1',c1,n1),('s2',c2,n2)):
            code = clean(code).upper()
            if not code or clean(name) or not valid_rep(code, name):
                continue
            existing = directory[kind].get(code)
            # A deleted or locked directory entry is never used to repair historical records.
            choices = names[kind,code]
            if (existing and (existing[2] or existing[3])) or len(choices)!=1:
                unresolved += 1
                continue
            fixes.append({'id':rid,'kind':kind,'code':code,'name':next(iter(choices))})
    token = hashlib.sha256(json.dumps(state,ensure_ascii=False,separators=(',',':')).encode()).hexdigest()
    return {'token':token,'newReps':new,'nameFixes':fixes,'blankCodes':blank_codes,
            'unresolvedNames':unresolved,'invalidCodes':sorted(invalid)}


def public_plan(p):
    return {'ok':True,'mode':'preview','token':p['token'],
            'missingS1':sum(r['kind']=='s1' for r in p['newReps']),
            'missingS2':sum(r['kind']=='s2' for r in p['newReps']),
            'fillableNames':len(p['nameFixes']), 'blankCodes':p['blankCodes'],
            'unresolvedNames':p['unresolvedNames'],'invalidCodes':len(p['invalidCodes']),
            'newRepSamples':p['newReps'][:30]}


def run(connect, lock, postgres, data_dir, mode='preview', options_json=''):
    if mode not in ('preview','apply'):
        return {'ok':False,'error':'Unknown data-check mode'}
    try:
        options = json.loads(options_json) if isinstance(options_json,str) and options_json else (options_json or {})
        if not isinstance(options,dict):
            raise ValueError()
        if mode=='apply' and (not options.get('token') or not all(isinstance(options.get(k),bool) for k in ('addReps','fillNames'))):
            raise ValueError()
    except (ValueError,TypeError):
        return {'ok':False,'error':'请重新检查数据，再选择修复内容。'}
    with lock:
        con = connect()
        try:
            if postgres:
                if mode=='preview':
                    con.execute('SET TRANSACTION ISOLATION LEVEL REPEATABLE READ READ ONLY')
                else:
                    con.execute("SET LOCAL lock_timeout='10s'")
                    con.execute('LOCK TABLE sales,reps,bcs_prefs IN SHARE ROW EXCLUSIVE MODE')
            elif mode=='apply':
                con.execute('BEGIN IMMEDIATE')
            state = snapshot(con)
            p = plan(state)
            if mode=='preview':
                return public_plan(p)
            if options['token'] != p['token']:
                return {'ok':False,'error':'数据已变化，请重新检查后再修复。'}
            adds = p['newReps'] if options['addReps'] else []
            fixes = p['nameFixes'] if options['fillNames'] else []
            if not adds and not fixes:
                return {'ok':True,'mode':'apply','added':0,'namesFilled':0,'backup':''}
            folder = Path(data_dir)/'repair-backups'
            folder.mkdir(parents=True,exist_ok=True,mode=0o700)
            folder.chmod(0o700)
            path = folder/('data-check-'+datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ')+'.json')
            fd = os.open(path,os.O_WRONLY|os.O_CREAT|os.O_EXCL,0o600)
            with os.fdopen(fd,'w',encoding='utf8') as f:
                json.dump({'token':p['token'],'beforeReps':state['reps'],'addReps':adds,'fillNames':fixes},f,ensure_ascii=False)
            for r in adds:
                con.execute('INSERT INTO reps(kind,code,name,group_,locked,deleted) VALUES (?,?,?,?,0,0)',(r['kind'],r['code'],r['name'],''))
            for kind, code_col, name_col in (('s1','rep1','rep1_name'),('s2','rep2','rep2_name')):
                con.executemany('UPDATE sales SET '+name_col+'=? WHERE id=? AND upper(trim('+code_col+'))=? AND trim(coalesce('+name_col+",''))=''",
                                [(r['name'],r['id'],r['code']) for r in fixes if r['kind']==kind])
            current = [tuple(r) for r in con.execute('SELECT kind,code,name,group_,locked,deleted FROM reps ORDER BY kind,code')]
            expected = sorted(state['reps']+[(r['kind'],r['code'],r['name'],'',0,0) for r in adds],key=lambda r:(r[0],r[1]))
            if current != expected:
                raise RuntimeError('分组保留检查未通过，修复已撤销。')
            import uuid
            con.execute('INSERT INTO bcs_prefs(key,value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value',('dataRevision',json.dumps(uuid.uuid4().hex)))
            con.commit()
            return {'ok':True,'mode':'apply','added':len(adds),'namesFilled':len(fixes),'backup':str(path)}
        finally:
            con.close()  # Any uncommitted transaction rolls back, including preview and failures.
