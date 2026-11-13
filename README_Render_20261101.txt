BCS 销售统计 V14.1 完整 Render 程序包（2026-11-01）

解压后文件直接位于项目根目录：server.py、requirements.txt、render.yaml、ui/。
把解压后的文件按目录结构更新到当前 Render 服务关联的 Git 仓库，保留原 Root Directory 设置。
不要把压缩包本身作为程序上传。若原项目位于仓库子目录，仍放在原子目录。
Build Command: pip install -r requirements.txt
Start Command: python server.py --no-browser
程序读取 Render 提供的 PORT。
保留原 DATABASE_URL、COST_VIEW_PASSWORD 和 AUTO_SYNC 设置。Render 应使用自己的数据库连接；本包没有 AWS 或 Neon 连接字符串。
提交后在现有服务部署该提交；首次打开 Ctrl+F5。
无需重新初始化、清空数据库或导入全部历史月份。

包含：真实销售工作表/表头识别；阻止表头或错位代码污染分组；导入管理月份倒序；固定单元格格式与快速启动；数据检查先预览再选择修复；IC 不要求 Sales2；客户动态按最新数据月统一排名、明细与趋势，删除/导入后清除旧缓存。
网页版备份明确为完整数据库 .db，上传还原覆盖销售、客户、库存、销售分组及设置。
本包不执行线上错误分组清理，不携带生产业务数据；此前 AWS 的清理结果不会自动复制到 Render。
Render 如仍有旧错误目录，需要在该数据库另行预览和清理。
包不包含 commission，不能替换佣金程序。
已完成隔离的本地接口/导入/还原测试；未在用户 Render 账户实际部署。
