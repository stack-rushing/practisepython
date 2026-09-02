# 学生随机点名系统（MySQL 版）

## 使用前准备

1. 安装并启动 MySQL Server。
2. 确认 `db_config.py` 中的主机、端口和用户名正确。
3. 在本文件夹打开 PowerShell，安装依赖：

   ```powershell
   py -m pip install -r requirements.txt
   ```

4. 启动程序：

   ```powershell
   py main.py
   ```

5. 根据提示输入 MySQL 密码。没有密码时直接按回车。

程序首次启动时会自动创建：

- 数据库：`student_roll_call`
- 数据表：`students`

学生姓名保存在 MySQL 中，关闭并重新运行程序后名单不会丢失。

如果出现“Access denied”，请检查用户名和密码。如果当前 MySQL 用户没有创建数据库的权限，请为其授权，或请管理员提前创建 `student_roll_call` 数据库。
