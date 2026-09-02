-- 创建学生随机点名系统使用的数据库
CREATE DATABASE IF NOT EXISTS student_roll_call
    CHARACTER SET utf8mb4
    COLLATE utf8mb4_unicode_ci;

USE student_roll_call;

-- 创建学生表
CREATE TABLE IF NOT EXISTS students (
    id INT PRIMARY KEY AUTO_INCREMENT,
    name VARCHAR(100) NOT NULL UNIQUE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 写入示例学生。重复执行此文件时不会重复添加。
INSERT IGNORE INTO students (name) VALUES
    ('张三'),
    ('李四'),
    ('王五'),
    ('赵六'),
    ('钱七');

-- 显示当前数据
SELECT id, name, created_at FROM students ORDER BY id;
