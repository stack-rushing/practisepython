import os
import random
import sqlite3
import sys
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook


APP_NAME = "学生随机点名系统"
APP_VERSION = "2.1.0"
TEMPLATE_NAME = "学生导入模板.xlsx"


def get_data_directory():
    """返回程序的数据保存目录，并确保目录存在。"""
    custom_directory = os.getenv("STUDENT_ROLL_CALL_DATA_DIR")
    local_app_data = os.getenv("LOCALAPPDATA")

    if custom_directory:
        data_directory = Path(custom_directory)
    elif local_app_data:
        data_directory = Path(local_app_data) / APP_NAME
    else:
        data_directory = Path.home() / APP_NAME

    data_directory.mkdir(parents=True, exist_ok=True)
    return data_directory


def resource_path(file_name):
    """返回资源路径，同时兼容源码运行和 PyInstaller。"""
    if hasattr(sys, "_MEIPASS"):
        base_path = Path(sys._MEIPASS)
    else:
        base_path = Path(__file__).resolve().parent
    return base_path / file_name


def masked_name(name):
    """点名滚动时只显示姓氏的第一个字，其余部分使用星号隐藏。"""
    clean_name = name.strip()
    if not clean_name:
        return "***"
    return f"{clean_name[0]}**"


class StudentDatabase:
    """使用 SQLite 保存学生名单。"""

    def __init__(self):
        self.database_path = get_data_directory() / "students.db"
        self.initialize()

    def connect(self):
        connection = sqlite3.connect(self.database_path)
        connection.row_factory = sqlite3.Row
        return connection

    def initialize(self):
        with self.connect() as connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS students (
                    student_no TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS app_meta (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                )
                """
            )

            seeded = connection.execute(
                "SELECT value FROM app_meta WHERE key = 'sample_data_created'"
            ).fetchone()
            if seeded is None:
                connection.executemany(
                    "INSERT OR IGNORE INTO students (student_no, name) VALUES (?, ?)",
                    [
                        ("S001", "张三"),
                        ("S002", "李四"),
                        ("S003", "王五"),
                        ("S004", "赵六"),
                        ("S005", "钱七"),
                    ],
                )
                connection.execute(
                    "INSERT INTO app_meta (key, value) VALUES (?, ?)",
                    ("sample_data_created", "1"),
                )

    def get_students(self):
        with self.connect() as connection:
            rows = connection.execute(
                "SELECT student_no, name FROM students ORDER BY student_no"
            ).fetchall()
        return [(row["student_no"], row["name"]) for row in rows]

    def replace_students(self, students):
        """用 Excel 中的有效学生整体替换当前名单。"""
        with self.connect() as connection:
            connection.execute("DELETE FROM students")
            connection.executemany(
                "INSERT INTO students (student_no, name) VALUES (?, ?)", students
            )


def excel_cell_text(value):
    """将 Excel 单元格内容转换为文字。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def find_excel_columns(worksheet):
    """在前 20 行中寻找学号和姓名表头。"""
    student_no_headers = {
        "学号", "学生学号", "studentid", "studentno", "studentnumber"
    }
    name_headers = {"姓名", "学生姓名", "name", "studentname"}

    for row_number, row in enumerate(
        worksheet.iter_rows(min_row=1, max_row=20, max_col=30), start=1
    ):
        student_no_column = None
        name_column = None

        for column_number, cell in enumerate(row, start=1):
            value = excel_cell_text(cell.value)
            normalized = value.lower().replace(" ", "").replace("_", "")
            if normalized in student_no_headers:
                student_no_column = column_number
            elif normalized in name_headers:
                name_column = column_number

        if student_no_column and name_column:
            return row_number, student_no_column, name_column

    return None


class StudentRollCallApp:
    """隐藏名单、带滚动悬念动画的学生随机点名界面。"""

    def __init__(self, root):
        self.root = root
        self.database = StudentDatabase()
        self.last_selected_student_no = None
        self.is_rolling = False
        self.rolling_students = []
        self.selected_student = None
        self.roll_step = 0
        self.total_roll_steps = 0
        self.color_index = 0

        self.root.title(f"{APP_NAME} {APP_VERSION}")
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = min(920, max(720, screen_width - 40))
        window_height = min(620, max(500, screen_height - 80))
        position_x = max(0, (screen_width - window_width) // 2)
        position_y = max(0, (screen_height - window_height) // 2)
        self.root.geometry(
            f"{window_width}x{window_height}+{position_x}+{position_y}"
        )
        self.root.minsize(min(760, window_width), min(520, window_height))
        self.root.configure(bg="#0F1E33")
        try:
            self.root.iconbitmap(default=str(resource_path("app_icon.ico")))
        except tk.TclError:
            pass

        self.count_var = tk.StringVar(value="已导入 0 名学生")
        self.stage_title_var = tk.StringVar(value="准备好了吗？")
        self.result_var = tk.StringVar(value="点击开始点名")
        self.detail_var = tk.StringVar(value="滚动时姓名会被隐藏，停止后揭晓结果")
        self.avoid_repeat_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="名单不会显示在界面上，保护点名悬念")

        self.configure_styles()
        self.create_widgets()
        self.refresh_count()

    def configure_styles(self):
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")

        style.configure("App.TFrame", background="#0F1E33")
        style.configure("Header.TLabel", font=("Microsoft YaHei UI", 23, "bold"),
                        foreground="#FFFFFF", background="#0F1E33")
        style.configure("HeaderSub.TLabel", font=("Microsoft YaHei UI", 10),
                        foreground="#A8BAD2", background="#0F1E33")
        style.configure("Count.TLabel", font=("Microsoft YaHei UI", 11, "bold"),
                        foreground="#FFD166", background="#0F1E33")
        style.configure("Stage.TFrame", background="#FFFFFF")
        style.configure("StageTitle.TLabel", font=("Microsoft YaHei UI", 15, "bold"),
                        foreground="#51647A", background="#FFFFFF")
        style.configure("Rolling.TLabel", font=("Microsoft YaHei UI", 48, "bold"),
                        foreground="#0E7A6D", background="#FFFFFF")
        style.configure("Detail.TLabel", font=("Microsoft YaHei UI", 12),
                        foreground="#607086", background="#FFFFFF")
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 15, "bold"),
                        padding=(38, 14))
        style.configure("Tool.TButton", font=("Microsoft YaHei UI", 10), padding=(15, 8))
        style.configure("Privacy.TLabel", font=("Microsoft YaHei UI", 9),
                        foreground="#8FA5BF", background="#0F1E33")
        style.configure("Status.TLabel", font=("Microsoft YaHei UI", 9),
                        foreground="#D4E0EF", background="#172B47")

    def create_widgets(self):
        main = ttk.Frame(self.root, style="App.TFrame", padding=26)
        main.pack(fill="both", expand=True)

        header = ttk.Frame(main, style="App.TFrame")
        header.pack(fill="x")
        ttk.Label(header, text=APP_NAME, style="Header.TLabel").pack(side="left")
        ttk.Label(header, textvariable=self.count_var, style="Count.TLabel").pack(
            side="right", pady=(10, 0)
        )
        ttk.Label(
            main,
            text="隐藏名单模式 · Excel 统一维护 · SQLite 本地保存",
            style="HeaderSub.TLabel",
        ).pack(anchor="w", pady=(4, 20))

        stage = ttk.Frame(main, style="Stage.TFrame", padding=(35, 28))
        stage.pack(fill="both", expand=True)
        ttk.Label(stage, textvariable=self.stage_title_var, style="StageTitle.TLabel").pack()
        self.result_label = ttk.Label(
            stage, textvariable=self.result_var, style="Rolling.TLabel", anchor="center"
        )
        self.result_label.pack(fill="x", pady=(34, 12))
        ttk.Label(stage, textvariable=self.detail_var, style="Detail.TLabel").pack()

        self.progress = ttk.Progressbar(stage, mode="indeterminate", length=430)
        self.progress.pack(pady=(24, 16))

        self.roll_button = ttk.Button(
            stage,
            text="开始随机点名",
            command=self.start_roll_call,
            style="Primary.TButton",
        )
        self.roll_button.pack(pady=(0, 12))
        ttk.Checkbutton(
            stage,
            text="避免连续抽中同一名学生",
            variable=self.avoid_repeat_var,
        ).pack()

        tools = ttk.Frame(main, style="App.TFrame")
        tools.pack(fill="x", pady=(18, 4))
        self.import_button = ttk.Button(
            tools,
            text="导入 / 更新 Excel 名单",
            command=self.import_excel,
            style="Tool.TButton",
        )
        self.import_button.pack(side="left")
        self.template_button = ttk.Button(
            tools,
            text="保存 Excel 导入模板",
            command=self.save_excel_template,
            style="Tool.TButton",
        )
        self.template_button.pack(side="left", padx=(10, 0))
        ttk.Label(
            tools,
            text="界面不展示学生名单；修改名单请编辑 Excel 后重新导入",
            style="Privacy.TLabel",
        ).pack(side="right", pady=(8, 0))

        ttk.Label(
            self.root, textvariable=self.status_var, style="Status.TLabel", padding=(14, 7)
        ).pack(side="bottom", fill="x")

    def refresh_count(self):
        count = len(self.database.get_students())
        self.count_var.set(f"已导入 {count} 名学生")
        return count

    def set_controls_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        self.roll_button.configure(state=state)
        self.import_button.configure(state=state)
        self.template_button.configure(state=state)

    def start_roll_call(self):
        if self.is_rolling:
            return

        students = self.database.get_students()
        if not students:
            messagebox.showinfo("无法点名", "学生名单为空，请先导入 Excel 名单。")
            return

        candidates = students
        if (
            self.avoid_repeat_var.get()
            and len(students) > 1
            and self.last_selected_student_no is not None
        ):
            candidates = [
                student for student in students
                if student[0] != self.last_selected_student_no
            ]

        self.is_rolling = True
        self.rolling_students = students
        self.selected_student = random.choice(candidates)
        self.roll_step = 0
        self.total_roll_steps = random.randint(32, 40)
        self.stage_title_var.set("正在随机抽取……")
        self.detail_var.set("姓名已隐藏，结果即将揭晓")
        self.status_var.set("点名进行中，请稍候……")
        self.set_controls_enabled(False)
        self.progress.start(10)
        self.run_roll_animation()

    def run_roll_animation(self):
        """快速滚动遮罩姓名，并在结束前逐渐减速。"""
        if self.roll_step >= self.total_roll_steps:
            self.finish_roll_call()
            return

        if self.roll_step == self.total_roll_steps - 1:
            rolling_student = self.selected_student
        else:
            rolling_student = random.choice(self.rolling_students)

        self.result_var.set(masked_name(rolling_student[1]))
        colors = ["#0E7A6D", "#2E75B6", "#D97706", "#7C3AED", "#C2415D"]
        self.color_index = (self.color_index + 1) % len(colors)
        ttk.Style().configure("Rolling.TLabel", foreground=colors[self.color_index])

        self.roll_step += 1
        progress = self.roll_step / self.total_roll_steps
        delay = 48 + int(300 * (progress ** 3))
        self.root.after(delay, self.run_roll_animation)

    def finish_roll_call(self):
        student_no, name = self.selected_student
        self.last_selected_student_no = student_no
        self.progress.stop()
        self.stage_title_var.set("本次点到")
        self.result_var.set(name)
        self.detail_var.set(f"学号：{student_no}")
        self.status_var.set(f"点名完成：{student_no}  {name}")
        self.roll_button.configure(text="再抽一次")
        self.is_rolling = False
        self.set_controls_enabled(True)
        self.flash_result(0)

    def flash_result(self, flash_count):
        """揭晓结果后短暂变色，增强停靠和揭晓感。"""
        if flash_count >= 6:
            ttk.Style().configure("Rolling.TLabel", foreground="#0E7A6D")
            return
        color = "#F59E0B" if flash_count % 2 == 0 else "#0E7A6D"
        ttk.Style().configure("Rolling.TLabel", foreground=color)
        self.root.after(180, lambda: self.flash_result(flash_count + 1))

    def import_excel(self):
        file_path = filedialog.askopenfilename(
            title="选择学生名单 Excel 文件",
            filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not file_path:
            return

        try:
            workbook = load_workbook(file_path, read_only=True, data_only=True)
            worksheet = workbook.active
            column_info = find_excel_columns(worksheet)
            if column_info is None:
                workbook.close()
                messagebox.showerror(
                    "无法识别表头",
                    "没有找到“学号”和“姓名”两列表头。\n\n"
                    "请使用程序提供的模板，或确认表头位于前 20 行。",
                )
                return

            header_row, student_no_column, name_column = column_info
            required_column_count = max(student_no_column, name_column)
            students = []
            seen_student_numbers = set()
            invalid_count = 0
            duplicate_count = 0

            for row in worksheet.iter_rows(
                min_row=header_row + 1, max_col=required_column_count
            ):
                student_no = excel_cell_text(row[student_no_column - 1].value)
                name = excel_cell_text(row[name_column - 1].value)
                if not student_no and not name:
                    continue
                if not student_no or not name:
                    invalid_count += 1
                    continue
                if student_no in seen_student_numbers:
                    duplicate_count += 1
                    continue
                seen_student_numbers.add(student_no)
                students.append((student_no, name))

            workbook.close()
            if not students:
                messagebox.showwarning("没有有效数据", "Excel 中没有可导入的有效学生。")
                return

            old_count = len(self.database.get_students())
            if old_count and not messagebox.askyesno(
                "确认更新名单",
                f"将用 Excel 中的 {len(students)} 名学生替换现有 {old_count} 名学生。\n\n"
                "是否继续？",
            ):
                return

            self.database.replace_students(students)
            self.last_selected_student_no = None
            self.refresh_count()
            self.stage_title_var.set("名单更新完成")
            self.result_var.set("准备开始点名")
            self.detail_var.set("滚动时仅显示姓氏和 **，停止后显示全名")
            self.status_var.set(f"Excel 导入成功：当前共 {len(students)} 名学生")
            messagebox.showinfo(
                "导入完成",
                f"当前名单：{len(students)} 名\n"
                f"重复学号：{duplicate_count} 条\n"
                f"缺少学号或姓名：{invalid_count} 条",
            )
        except Exception as error:
            messagebox.showerror("导入失败", f"无法读取该 Excel 文件：\n{error}")

    def save_excel_template(self):
        template_path = resource_path(TEMPLATE_NAME)
        if not template_path.exists():
            messagebox.showerror("模板不存在", "安装文件中没有找到 Excel 导入模板。")
            return

        destination = filedialog.asksaveasfilename(
            title="保存学生导入模板",
            initialfile=TEMPLATE_NAME,
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if not destination:
            return

        try:
            Path(destination).write_bytes(template_path.read_bytes())
            self.status_var.set(f"模板已保存到：{destination}")
            messagebox.showinfo("保存成功", "Excel 导入模板已保存。")
        except OSError as error:
            messagebox.showerror("保存失败", f"无法保存模板：\n{error}")


def main():
    root = tk.Tk()
    StudentRollCallApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
