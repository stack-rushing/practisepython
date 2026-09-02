import random
from getpass import getpass

import mysql.connector
from mysql.connector import Error, IntegrityError

from db_config import DATABASE_NAME, MYSQL_CONFIG


# 程序启动时读取一次密码，后续数据库操作会使用该密码。
mysql_password = ""


def get_connection(include_database=True):
    """创建并返回一个 MySQL 连接。"""
    connection_config = MYSQL_CONFIG.copy()
    connection_config["password"] = mysql_password

    if include_database:
        connection_config["database"] = DATABASE_NAME

    return mysql.connector.connect(**connection_config)


def initialize_database():
    """首次运行时自动创建数据库和学生表。"""
    connection = get_connection(include_database=False)
    cursor = connection.cursor()

    try:
        # 数据库名称来自 db_config.py，不接收用户输入。
        cursor.execute(
            f"CREATE DATABASE IF NOT EXISTS `{DATABASE_NAME}` "
            "CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
        )
    finally:
        cursor.close()
        connection.close()

    connection = get_connection()
    cursor = connection.cursor()

    try:
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS students (
                id INT PRIMARY KEY AUTO_INCREMENT,
                name VARCHAR(100) NOT NULL UNIQUE,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        connection.commit()
    finally:
        cursor.close()
        connection.close()


def input_student_name(message):
    """读取非空的学生姓名。"""
    while True:
        name = input(message).strip()
        if name:
            return name
        print("输入不能为空，请重新输入。")


def add_student():
    """向 MySQL 学生表中添加一名学生。"""
    name = input_student_name("请输入要添加的学生姓名：")
    connection = None
    cursor = None

    try:
        connection = get_connection()
        cursor = connection.cursor()
        cursor.execute("INSERT INTO students (name) VALUES (%s)", (name,))
        connection.commit()
        print(f"学生“{name}”添加成功。")
    except IntegrityError:
        print("该学生已存在。")
    except Error as error:
        print(f"添加失败：{error}")
    finally:
        if cursor is not None:
            cursor.close()
        if connection is not None and connection.is_connected():
            connection.close()


def delete_student():
    """根据姓名删除学生。"""
    name = input_student_name("请输入要删除的学生姓名：")
    connection = None
    cursor = None

    try:
        connection = get_connection()
        cursor = connection.cursor()
        cursor.execute("DELETE FROM students WHERE name = %s", (name,))
        connection.commit()

        if cursor.rowcount == 0:
            print("未找到该学生。")
        else:
            print(f"学生“{name}”删除成功。")
    except Error as error:
        print(f"删除失败：{error}")
    finally:
        if cursor is not None:
            cursor.close()
        if connection is not None and connection.is_connected():
            connection.close()


def random_student():
    """使用 random 模块随机抽取一名学生。"""
    connection = None
    cursor = None

    try:
        connection = get_connection()
        cursor = connection.cursor()
        cursor.execute("SELECT name FROM students")
        rows = cursor.fetchall()

        if not rows:
            print("学生名单为空，无法点名。")
            return

        # 查询结果中的每一行都是元组，例如：(“张三”,)
        student_names = [row[0] for row in rows]
        selected_student = random.choice(student_names)
        print(f"本次点到的学生是：{selected_student}")
    except Error as error:
        print(f"点名失败：{error}")
    finally:
        if cursor is not None:
            cursor.close()
        if connection is not None and connection.is_connected():
            connection.close()


def show_students():
    """按照添加顺序显示全部学生和总人数。"""
    connection = None
    cursor = None

    try:
        connection = get_connection()
        cursor = connection.cursor()
        cursor.execute("SELECT name FROM students ORDER BY id")
        rows = cursor.fetchall()

        if not rows:
            print("当前学生名单为空。")
            return

        print("\n===== 当前学生名单 =====")
        for number, row in enumerate(rows, start=1):
            print(f"{number}. {row[0]}")
        print(f"学生总人数：{len(rows)}")
    except Error as error:
        print(f"查询失败：{error}")
    finally:
        if cursor is not None:
            cursor.close()
        if connection is not None and connection.is_connected():
            connection.close()


def show_menu():
    """显示程序主菜单。"""
    print("\n===== 学生随机点名系统 =====")
    print("1. 添加学生")
    print("2. 删除学生")
    print("3. 随机抽取学生")
    print("4. 查看全部学生")
    print("5. 退出程序")
    print("==========================")


def main():
    """程序入口，负责初始化数据库并循环处理菜单选择。"""
    global mysql_password

    print("正在连接 MySQL……")
    mysql_password = getpass("请输入 MySQL 密码（没有密码可直接回车）：")

    try:
        initialize_database()
    except Error as error:
        print(f"无法初始化数据库：{error}")
        print("请确认 MySQL 已启动，并检查 db_config.py 中的连接配置。")
        return

    print(f"数据库 `{DATABASE_NAME}` 已准备就绪。")

    while True:
        show_menu()
        choice = input("请输入功能序号：").strip()

        if choice == "1":
            add_student()
        elif choice == "2":
            delete_student()
        elif choice == "3":
            random_student()
        elif choice == "4":
            show_students()
        elif choice == "5":
            print("程序已退出，再见！")
            break
        else:
            print("输入错误，请输入 1～5 之间的数字。")


if __name__ == "__main__":
    main()
