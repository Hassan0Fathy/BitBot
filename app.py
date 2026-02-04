import calendar
import sqlite3
from datetime import date
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from openpyxl import Workbook


DB_PATH = "employees.db"


class Database:
    def __init__(self, path=DB_PATH):
        self.path = path
        self._init_db()

    def _connect(self):
        return sqlite3.connect(self.path)

    def _init_db(self):
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS employees (
                    emp_no TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    salary REAL NOT NULL
                )
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS attendance (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    emp_no TEXT NOT NULL,
                    day TEXT NOT NULL,
                    status TEXT NOT NULL,
                    UNIQUE(emp_no, day),
                    FOREIGN KEY(emp_no) REFERENCES employees(emp_no)
                )
                """
            )
            conn.commit()

    def add_employee(self, emp_no, name, salary):
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO employees(emp_no, name, salary) VALUES (?, ?, ?)",
                (emp_no, name, salary),
            )
            conn.commit()

    def update_employee(self, emp_no, name, salary):
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "UPDATE employees SET name = ?, salary = ? WHERE emp_no = ?",
                (name, salary, emp_no),
            )
            conn.commit()

    def delete_employee(self, emp_no):
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM employees WHERE emp_no = ?", (emp_no,))
            cursor.execute("DELETE FROM attendance WHERE emp_no = ?", (emp_no,))
            conn.commit()

    def list_employees(self):
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT emp_no, name, salary FROM employees ORDER BY emp_no")
            return cursor.fetchall()

    def get_employee(self, emp_no):
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT emp_no, name, salary FROM employees WHERE emp_no = ?",
                (emp_no,),
            )
            return cursor.fetchone()

    def upsert_attendance(self, emp_no, day, status):
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                INSERT INTO attendance(emp_no, day, status)
                VALUES (?, ?, ?)
                ON CONFLICT(emp_no, day)
                DO UPDATE SET status = excluded.status
                """,
                (emp_no, day, status),
            )
            conn.commit()

    def list_attendance(self, emp_no):
        with self._connect() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT day, status FROM attendance WHERE emp_no = ? ORDER BY day DESC",
                (emp_no,),
            )
            return cursor.fetchall()

    def attendance_summary(self, emp_no, month, year):
        with self._connect() as conn:
            cursor = conn.cursor()
            start_date = f"{year:04d}-{month:02d}-01"
            end_date = f"{year:04d}-{month:02d}-{calendar.monthrange(year, month)[1]:02d}"
            cursor.execute(
                """
                SELECT
                    SUM(CASE WHEN status = 'حاضر' THEN 1 ELSE 0 END) AS present_count,
                    SUM(CASE WHEN status = 'غائب' THEN 1 ELSE 0 END) AS absent_count
                FROM attendance
                WHERE emp_no = ? AND day BETWEEN ? AND ?
                """,
                (emp_no, start_date, end_date),
            )
            row = cursor.fetchone()
            return row[0] or 0, row[1] or 0


class EmployeeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("إدارة الموظفين والحضور والمرتبات")
        self.root.geometry("980x620")
        self.root.configure(bg="#f5f5f5")

        self.db = Database()

        self._build_ui()
        self.refresh_employees()

    def _build_ui(self):
        header = tk.Label(
            self.root,
            text="نظام إدارة الموظفين والحضور والمرتبات",
            font=("Arial", 18, "bold"),
            bg="#f5f5f5",
        )
        header.pack(pady=10)

        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.employee_frame = ttk.Frame(notebook)
        self.attendance_frame = ttk.Frame(notebook)
        self.salary_frame = ttk.Frame(notebook)

        notebook.add(self.employee_frame, text="الموظفين")
        notebook.add(self.attendance_frame, text="الحضور والغياب")
        notebook.add(self.salary_frame, text="المرتبات")

        self._build_employee_tab()
        self._build_attendance_tab()
        self._build_salary_tab()

    def _build_employee_tab(self):
        form = ttk.LabelFrame(self.employee_frame, text="بيانات الموظف")
        form.pack(fill="x", padx=10, pady=10)

        ttk.Label(form, text="رقم الموظف:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        ttk.Label(form, text="اسم الموظف:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        ttk.Label(form, text="المرتب الشهري:").grid(row=0, column=4, padx=5, pady=5, sticky="e")

        self.emp_no_var = tk.StringVar()
        self.emp_name_var = tk.StringVar()
        self.emp_salary_var = tk.StringVar()

        ttk.Entry(form, textvariable=self.emp_no_var, width=20, justify="right").grid(
            row=0, column=1, padx=5, pady=5
        )
        ttk.Entry(form, textvariable=self.emp_name_var, width=30, justify="right").grid(
            row=0, column=3, padx=5, pady=5
        )
        ttk.Entry(form, textvariable=self.emp_salary_var, width=20, justify="right").grid(
            row=0, column=5, padx=5, pady=5
        )

        actions = ttk.Frame(form)
        actions.grid(row=1, column=0, columnspan=6, pady=10)

        ttk.Button(actions, text="إضافة", command=self.add_employee).grid(
            row=0, column=0, padx=5
        )
        ttk.Button(actions, text="تعديل", command=self.update_employee).grid(
            row=0, column=1, padx=5
        )
        ttk.Button(actions, text="حذف", command=self.delete_employee).grid(
            row=0, column=2, padx=5
        )
        ttk.Button(actions, text="تفريغ", command=self.clear_employee_form).grid(
            row=0, column=3, padx=5
        )

        list_frame = ttk.LabelFrame(self.employee_frame, text="قائمة الموظفين")
        list_frame.pack(fill="both", expand=True, padx=10, pady=10)

        columns = ("emp_no", "name", "salary")
        self.employee_tree = ttk.Treeview(
            list_frame, columns=columns, show="headings", height=12
        )
        self.employee_tree.heading("emp_no", text="رقم الموظف")
        self.employee_tree.heading("name", text="اسم الموظف")
        self.employee_tree.heading("salary", text="المرتب الشهري")
        self.employee_tree.column("emp_no", width=120, anchor="center")
        self.employee_tree.column("name", width=300, anchor="center")
        self.employee_tree.column("salary", width=150, anchor="center")
        self.employee_tree.pack(fill="both", expand=True)

        self.employee_tree.bind("<<TreeviewSelect>>", self.on_employee_select)

    def _build_attendance_tab(self):
        form = ttk.LabelFrame(self.attendance_frame, text="تسجيل الحضور والغياب")
        form.pack(fill="x", padx=10, pady=10)

        ttk.Label(form, text="الموظف:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        ttk.Label(form, text="التاريخ:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        ttk.Label(form, text="الحالة:").grid(row=0, column=4, padx=5, pady=5, sticky="e")

        self.att_emp_var = tk.StringVar()
        self.att_date_var = tk.StringVar(value=date.today().isoformat())
        self.att_status_var = tk.StringVar(value="حاضر")

        self.att_emp_combo = ttk.Combobox(
            form, textvariable=self.att_emp_var, state="readonly", width=30
        )
        self.att_emp_combo.grid(row=0, column=1, padx=5, pady=5)
        self.att_emp_combo.bind("<<ComboboxSelected>>", self.refresh_attendance_list)

        ttk.Entry(form, textvariable=self.att_date_var, width=20, justify="right").grid(
            row=0, column=3, padx=5, pady=5
        )
        ttk.Combobox(
            form,
            textvariable=self.att_status_var,
            values=["حاضر", "غائب"],
            state="readonly",
            width=15,
        ).grid(row=0, column=5, padx=5, pady=5)

        ttk.Button(form, text="حفظ", command=self.save_attendance).grid(
            row=1, column=0, columnspan=6, pady=10
        )

        list_frame = ttk.LabelFrame(self.attendance_frame, text="سجل الحضور")
        list_frame.pack(fill="both", expand=True, padx=10, pady=10)

        columns = ("day", "status")
        self.attendance_tree = ttk.Treeview(
            list_frame, columns=columns, show="headings", height=12
        )
        self.attendance_tree.heading("day", text="التاريخ")
        self.attendance_tree.heading("status", text="الحالة")
        self.attendance_tree.column("day", width=150, anchor="center")
        self.attendance_tree.column("status", width=100, anchor="center")
        self.attendance_tree.pack(fill="both", expand=True)

    def _build_salary_tab(self):
        form = ttk.LabelFrame(self.salary_frame, text="حساب المرتبات")
        form.pack(fill="x", padx=10, pady=10)

        ttk.Label(form, text="الموظف:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        ttk.Label(form, text="الشهر:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        ttk.Label(form, text="السنة:").grid(row=0, column=4, padx=5, pady=5, sticky="e")

        self.salary_emp_var = tk.StringVar()
        self.salary_month_var = tk.IntVar(value=date.today().month)
        self.salary_year_var = tk.IntVar(value=date.today().year)

        self.salary_emp_combo = ttk.Combobox(
            form, textvariable=self.salary_emp_var, state="readonly", width=30
        )
        self.salary_emp_combo.grid(row=0, column=1, padx=5, pady=5)

        month_values = list(range(1, 13))
        ttk.Combobox(
            form,
            textvariable=self.salary_month_var,
            values=month_values,
            state="readonly",
            width=10,
        ).grid(row=0, column=3, padx=5, pady=5)

        year_values = list(range(date.today().year - 5, date.today().year + 6))
        ttk.Combobox(
            form,
            textvariable=self.salary_year_var,
            values=year_values,
            state="readonly",
            width=10,
        ).grid(row=0, column=5, padx=5, pady=5)

        ttk.Button(form, text="احتساب", command=self.calculate_salary).grid(
            row=1, column=0, columnspan=6, pady=10
        )

        summary_frame = ttk.LabelFrame(self.salary_frame, text="نتائج المرتب")
        summary_frame.pack(fill="x", padx=10, pady=10)

        self.total_days_var = tk.StringVar(value="0")
        self.present_days_var = tk.StringVar(value="0")
        self.absent_days_var = tk.StringVar(value="0")
        self.net_salary_var = tk.StringVar(value="0")

        ttk.Label(summary_frame, text="عدد أيام العمل:").grid(
            row=0, column=0, padx=5, pady=5, sticky="e"
        )
        ttk.Label(summary_frame, textvariable=self.total_days_var).grid(
            row=0, column=1, padx=5, pady=5
        )
        ttk.Label(summary_frame, text="عدد أيام الحضور:").grid(
            row=0, column=2, padx=5, pady=5, sticky="e"
        )
        ttk.Label(summary_frame, textvariable=self.present_days_var).grid(
            row=0, column=3, padx=5, pady=5
        )
        ttk.Label(summary_frame, text="عدد أيام الغياب:").grid(
            row=1, column=0, padx=5, pady=5, sticky="e"
        )
        ttk.Label(summary_frame, textvariable=self.absent_days_var).grid(
            row=1, column=1, padx=5, pady=5
        )
        ttk.Label(summary_frame, text="صافي المرتب المستحق:").grid(
            row=1, column=2, padx=5, pady=5, sticky="e"
        )
        ttk.Label(summary_frame, textvariable=self.net_salary_var).grid(
            row=1, column=3, padx=5, pady=5
        )

        ttk.Button(
            self.salary_frame, text="تصدير إلى Excel", command=self.export_salary
        ).pack(pady=10)

    def refresh_employees(self):
        for item in self.employee_tree.get_children():
            self.employee_tree.delete(item)
        for emp_no, name, salary in self.db.list_employees():
            self.employee_tree.insert("", "end", values=(emp_no, name, salary))
        self._refresh_employee_comboboxes()

    def _refresh_employee_comboboxes(self):
        employees = self.db.list_employees()
        display = [f"{emp_no} - {name}" for emp_no, name, _ in employees]
        self.att_emp_combo["values"] = display
        self.salary_emp_combo["values"] = display

    def _parse_employee_selection(self, selection):
        if not selection:
            return None
        return selection.split("-")[0].strip()

    def add_employee(self):
        emp_no = self.emp_no_var.get().strip()
        name = self.emp_name_var.get().strip()
        salary_text = self.emp_salary_var.get().strip()
        if not emp_no or not name or not salary_text:
            messagebox.showwarning("تنبيه", "يرجى إدخال جميع البيانات المطلوبة.")
            return
        try:
            salary = float(salary_text)
        except ValueError:
            messagebox.showerror("خطأ", "المرتب يجب أن يكون رقمًا صحيحًا أو عشريًا.")
            return
        try:
            self.db.add_employee(emp_no, name, salary)
        except sqlite3.IntegrityError:
            messagebox.showerror("خطأ", "رقم الموظف موجود مسبقًا.")
            return
        self.refresh_employees()
        self.clear_employee_form()
        messagebox.showinfo("تم", "تمت إضافة الموظف بنجاح.")

    def update_employee(self):
        emp_no = self.emp_no_var.get().strip()
        name = self.emp_name_var.get().strip()
        salary_text = self.emp_salary_var.get().strip()
        if not emp_no or not name or not salary_text:
            messagebox.showwarning("تنبيه", "يرجى اختيار موظف وتعديل البيانات.")
            return
        try:
            salary = float(salary_text)
        except ValueError:
            messagebox.showerror("خطأ", "المرتب يجب أن يكون رقمًا صحيحًا أو عشريًا.")
            return
        self.db.update_employee(emp_no, name, salary)
        self.refresh_employees()
        messagebox.showinfo("تم", "تم تعديل بيانات الموظف.")

    def delete_employee(self):
        emp_no = self.emp_no_var.get().strip()
        if not emp_no:
            messagebox.showwarning("تنبيه", "يرجى اختيار موظف للحذف.")
            return
        if not messagebox.askyesno("تأكيد", "هل أنت متأكد من حذف الموظف؟"):
            return
        self.db.delete_employee(emp_no)
        self.refresh_employees()
        self.clear_employee_form()
        messagebox.showinfo("تم", "تم حذف الموظف.")

    def clear_employee_form(self):
        self.emp_no_var.set("")
        self.emp_name_var.set("")
        self.emp_salary_var.set("")

    def on_employee_select(self, event):
        selected = self.employee_tree.selection()
        if not selected:
            return
        values = self.employee_tree.item(selected[0], "values")
        self.emp_no_var.set(values[0])
        self.emp_name_var.set(values[1])
        self.emp_salary_var.set(values[2])

    def save_attendance(self):
        selection = self.att_emp_var.get()
        emp_no = self._parse_employee_selection(selection)
        day = self.att_date_var.get().strip()
        status = self.att_status_var.get().strip()
        if not emp_no or not day or not status:
            messagebox.showwarning("تنبيه", "يرجى اختيار الموظف وإدخال التاريخ والحالة.")
            return
        try:
            date.fromisoformat(day)
        except ValueError:
            messagebox.showerror("خطأ", "صيغة التاريخ يجب أن تكون YYYY-MM-DD.")
            return
        self.db.upsert_attendance(emp_no, day, status)
        self.refresh_attendance_list()
        messagebox.showinfo("تم", "تم حفظ سجل الحضور.")

    def refresh_attendance_list(self, event=None):
        for item in self.attendance_tree.get_children():
            self.attendance_tree.delete(item)
        selection = self.att_emp_var.get()
        emp_no = self._parse_employee_selection(selection)
        if not emp_no:
            return
        for day, status in self.db.list_attendance(emp_no):
            self.attendance_tree.insert("", "end", values=(day, status))

    def calculate_salary(self):
        selection = self.salary_emp_var.get()
        emp_no = self._parse_employee_selection(selection)
        if not emp_no:
            messagebox.showwarning("تنبيه", "يرجى اختيار الموظف.")
            return
        employee = self.db.get_employee(emp_no)
        if not employee:
            messagebox.showerror("خطأ", "الموظف غير موجود.")
            return
        _, _, salary = employee
        month = int(self.salary_month_var.get())
        year = int(self.salary_year_var.get())
        total_days = calendar.monthrange(year, month)[1]
        present_days, absent_days = self.db.attendance_summary(emp_no, month, year)
        daily_rate = salary / total_days if total_days else 0
        net_salary = salary - (daily_rate * absent_days)

        self.total_days_var.set(str(total_days))
        self.present_days_var.set(str(present_days))
        self.absent_days_var.set(str(absent_days))
        self.net_salary_var.set(f"{net_salary:.2f}")

    def export_salary(self):
        selection = self.salary_emp_var.get()
        emp_no = self._parse_employee_selection(selection)
        if not emp_no:
            messagebox.showwarning("تنبيه", "يرجى اختيار الموظف لتصدير التقرير.")
            return
        employee = self.db.get_employee(emp_no)
        if not employee:
            messagebox.showerror("خطأ", "الموظف غير موجود.")
            return
        month = int(self.salary_month_var.get())
        year = int(self.salary_year_var.get())
        total_days = calendar.monthrange(year, month)[1]
        present_days, absent_days = self.db.attendance_summary(emp_no, month, year)
        daily_rate = employee[2] / total_days if total_days else 0
        net_salary = employee[2] - (daily_rate * absent_days)

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="حفظ تقرير المرتب",
        )
        if not file_path:
            return

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "تقرير المرتب"

        sheet.append(["رقم الموظف", employee[0]])
        sheet.append(["اسم الموظف", employee[1]])
        sheet.append(["الشهر", month])
        sheet.append(["السنة", year])
        sheet.append(["عدد أيام العمل", total_days])
        sheet.append(["عدد أيام الحضور", present_days])
        sheet.append(["عدد أيام الغياب", absent_days])
        sheet.append(["صافي المرتب المستحق", f"{net_salary:.2f}"])

        workbook.save(file_path)
        messagebox.showinfo("تم", "تم تصدير تقرير المرتب بنجاح.")


if __name__ == "__main__":
    root = tk.Tk()
    app = EmployeeApp(root)
    root.mainloop()
