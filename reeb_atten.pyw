import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import csv

# Database Setup
def init_db():
    conn = sqlite3.connect("attendance.db")
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS employees (
        employee_number INTEGER PRIMARY KEY,
        first_name TEXT NOT NULL,
        last_name TEXT NOT NULL,
        department TEXT NOT NULL,
        hire_date DATE NOT NULL,
        termination_date DATE)''')
    c.execute('''CREATE TABLE IF NOT EXISTS attendance (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_number INTEGER NOT NULL,
        date DATE NOT NULL,
        present BOOLEAN DEFAULT 0,
        tardy BOOLEAN DEFAULT 0,
        early_out BOOLEAN DEFAULT 0,
        absent BOOLEAN DEFAULT 0,
        bonus_points BOOLEAN DEFAULT 0,
        notes TEXT,
        FOREIGN KEY (employee_number) REFERENCES employees(employee_number),
        UNIQUE (employee_number, date))''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_attendance_date ON attendance(date)')
    conn.commit()
    conn.close()

# GUI Class
class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Reeb Exterior Employee Attendance System")
        self.conn = sqlite3.connect("attendance.db")
        
        # Main Frame
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Buttons for actions
        ttk.Button(self.main_frame, text="Add Employee", command=self.add_employee).grid(row=0, column=0, pady=5)
        ttk.Button(self.main_frame, text="Record Attendance", command=self.record_attendance).grid(row=0, column=1, pady=5)
        ttk.Button(self.main_frame, text="Daily Report", command=self.daily_report).grid(row=0, column=2, pady=5)
        ttk.Button(self.main_frame, text="Weekly Summary", command=self.weekly_summary).grid(row=0, column=3, pady=5)
        ttk.Button(self.main_frame, text="Monthly Summary", command=self.monthly_summary).grid(row=0, column=4, pady=5)
        ttk.Button(self.main_frame, text="Import Names", command=self.import_names).grid(row=0, column=5, pady=5)

    def add_employee(self):
        # Popup window for adding employee
        win = tk.Toplevel(self.root)
        win.title("Add Employee")
        win.geometry("300x200")  # Set a reasonable size for the window
    
        # Labels and Entry Fields
        ttk.Label(win, text="Employee Number").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        emp_num = ttk.Entry(win)
        emp_num.grid(row=0, column=1, padx=5, pady=5)
        emp_num.focus()  # Set focus to Employee Number field
        
        ttk.Label(win, text="First Name").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        fname = ttk.Entry(win)
        fname.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(win, text="Last Name").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        lname = ttk.Entry(win)
        lname.grid(row=2, column=1, padx=5, pady=5)
        
        # Department Dropdown
        ttk.Label(win, text="Department").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        departments = ["PREP", "ACRO", "RKG", "SDL", "EPL"]  # Predefined list
        dept = ttk.Combobox(win, values=departments, state="readonly")  # Dropdown with options
        dept.grid(row=3, column=1, padx=5, pady=5)
        #dept.set("PREP")  # Default selection

        ttk.Label(win, text="Hire Date").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
        hdate = ttk.Entry(win)
        hdate.grid(row=4, column=1, padx=5, pady=5)
        
        def save_employee():
            # Get values from entry fields
            emp_num_val = emp_num.get().strip()
            fname_val = fname.get().strip()
            lname_val = lname.get().strip()
            dept_val = dept.get()
            hdate_val = hdate.get().strip()

            # Validation
            if not emp_num_val or not fname_val or not lname_val or not dept_val or not hdate_val:
                messagebox.showerror("Error", "All fields are required!")
                return
            
            try:
                emp_num_int = int(emp_num_val)  # Ensure employee_number is an integer
            except ValueError:
                messagebox.showerror("Error", "Employee Number must be an integer!")
                return
            
            # Database operation
            try:
                c = self.conn.cursor()
                # Insert new employee with current date as hire_date
                c.execute("""
                    INSERT INTO employees (employee_number, first_name, last_name, department, hire_date)
                    VALUES (?, ?, ?, ?, ?)""",
                    (emp_num_int, fname_val, lname_val, dept_val, hdate_val))
                self.conn.commit()
                
                # Success feedback
                messagebox.showinfo("Success", f"Employee {fname_val} {lname_val} (#{emp_num_int}) added!")
                
                # Clear the form
                emp_num.delete(0, tk.END)
                fname.delete(0, tk.END)
                lname.delete(0, tk.END)
                dept.delete(0, tk.END)
                hdate.delete(0, tk.END)
                
                # Optional: Close the window after saving
                # win.destroy()
                
            except sqlite3.IntegrityError:
                # Handle duplicate employee_number (PRIMARY KEY violation)
                messagebox.showerror("Error", f"Employee Number {emp_num_val} already exists!")
            except Exception as e:
                # Catch any other unexpected errors
                messagebox.showerror("Error", f"An error occurred: {str(e)}")
        
        # Save Button
        ttk.Button(win, text="Save", command=save_employee).grid(row=5, column=1, padx=5, pady=10)
        
        # Optional: Add a Cancel button to close the window without saving
        ttk.Button(win, text="Cancel", command=win.destroy).grid(row=5, column=0, padx=5, pady=10)        
        
    def record_attendance(self):
        win = tk.Toplevel(self.root)
        win.title("Record Attendance")
        win.geometry("800x600")

        # Current date
        current_date = datetime.now().strftime("%Y-%m-%d")
        
        # Date entry
        ttk.Label(win, text="Date").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        date_entry = ttk.Entry(win)
        date_entry.insert(0, current_date)
        date_entry.grid(row=0, column=1, padx=5, pady=5)

        # Canvas with scrollbar
        canvas = tk.Canvas(win)
        scrollbar = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Get employees, sorted by last_name then department
        c = self.conn.cursor()
        c.execute("""
            SELECT employee_number, first_name, last_name, department 
            FROM employees 
            WHERE termination_date IS NULL 
            ORDER BY last_name, department
        """)
        employees = c.fetchall()

        # Attendance variables storage
        attendance_vars = {}
        notes_entries = {}

        # Select All variables
        select_all_vars = {
            'present': tk.BooleanVar(),
            'tardy': tk.BooleanVar(),
            'early': tk.BooleanVar(),
            'absent': tk.BooleanVar(),
            'bonus': tk.BooleanVar()
        }

        # Header with Select All
        ttk.Label(scrollable_frame, text="Employee").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Checkbutton(scrollable_frame, text="All Present", variable=select_all_vars['present'], 
                    command=lambda: toggle_all('present')).grid(row=0, column=1, padx=5, pady=5)
        ttk.Checkbutton(scrollable_frame, text="All Tardy", variable=select_all_vars['tardy'],
                    command=lambda: toggle_all('tardy')).grid(row=0, column=2, padx=5, pady=5)
        ttk.Checkbutton(scrollable_frame, text="All Early Out", variable=select_all_vars['early'],
                    command=lambda: toggle_all('early')).grid(row=0, column=3, padx=5, pady=5)
        ttk.Checkbutton(scrollable_frame, text="All Absent", variable=select_all_vars['absent'],
                    command=lambda: toggle_all('absent')).grid(row=0, column=4, padx=5, pady=5)
        ttk.Checkbutton(scrollable_frame, text="All Bonus", variable=select_all_vars['bonus'],
                    command=lambda: toggle_all('bonus')).grid(row=0, column=5, padx=5, pady=5)
        ttk.Label(scrollable_frame, text="Notes").grid(row=0, column=6, padx=5, pady=5)

        # Create employee rows
        for i, emp in enumerate(employees, 1):
            emp_num = emp[0]
            name_label = ttk.Label(scrollable_frame, text=f"{emp[1]} {emp[2]} (#{emp_num}) - {emp[3]}")
            name_label.grid(row=i, column=0, padx=5, pady=2, sticky=tk.W)

            # Create variables for each checkbox
            vars_dict = {
                'present': tk.BooleanVar(),
                'tardy': tk.BooleanVar(),
                'early': tk.BooleanVar(),
                'absent': tk.BooleanVar(),
                'bonus': tk.BooleanVar()
            }
            attendance_vars[emp_num] = vars_dict

            # Checkboxes
            ttk.Checkbutton(scrollable_frame, variable=vars_dict['present']).grid(row=i, column=1, padx=5, pady=2)
            ttk.Checkbutton(scrollable_frame, text="", variable=vars_dict['tardy']).grid(row=i, column=2, padx=5, pady=2)
            ttk.Checkbutton(scrollable_frame, text="", variable=vars_dict['early']).grid(row=i, column=3, padx=5, pady=2)
            ttk.Checkbutton(scrollable_frame, text="", variable=vars_dict['absent']).grid(row=i, column=4, padx=5, pady=2)
            ttk.Checkbutton(scrollable_frame, text="", variable=vars_dict['bonus']).grid(row=i, column=5, padx=5, pady=2)

            # Notes entry
            notes_entry = ttk.Entry(scrollable_frame, width=20)
            notes_entry.grid(row=i, column=6, padx=5, pady=2)
            notes_entries[emp_num] = notes_entry

        def toggle_all(category):
            state = select_all_vars[category].get()
            for emp_num in attendance_vars:
                attendance_vars[emp_num][category].set(state)

        def load_existing():
            date = date_entry.get()
            c.execute("SELECT employee_number, present, tardy, early_out, absent, bonus_points, notes FROM attendance WHERE date = ?", (date,))
            records = {row[0]: row[1:] for row in c.fetchall()}
            
            for emp_num in attendance_vars:
                if emp_num in records:
                    record = records[emp_num]
                    attendance_vars[emp_num]['present'].set(record[0])
                    attendance_vars[emp_num]['tardy'].set(record[1])
                    attendance_vars[emp_num]['early'].set(record[2])
                    attendance_vars[emp_num]['absent'].set(record[3])
                    attendance_vars[emp_num]['bonus'].set(record[4])
                    notes_entries[emp_num].delete(0, tk.END)
                    notes_entries[emp_num].insert(0, record[5] or "")
                else:
                    # Clear if no record exists
                    for var in attendance_vars[emp_num].values():
                        var.set(False)
                    notes_entries[emp_num].delete(0, tk.END)
            
            # Update select all checkboxes based on current state
            for category in select_all_vars:
                all_checked = all(attendance_vars[emp_num][category].get() for emp_num in attendance_vars)
                select_all_vars[category].set(all_checked)

        def save_attendance():
            date = date_entry.get()
            c = self.conn.cursor()
            try:
                for emp_num in attendance_vars:
                    c.execute("""
                        INSERT OR REPLACE INTO attendance 
                        (employee_number, date, present, tardy, early_out, absent, bonus_points, notes)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                        (emp_num, date,
                        attendance_vars[emp_num]['present'].get(),
                        attendance_vars[emp_num]['tardy'].get(),
                        attendance_vars[emp_num]['early'].get(),
                        attendance_vars[emp_num]['absent'].get(),
                        attendance_vars[emp_num]['bonus'].get(),
                        notes_entries[emp_num].get()))
                self.conn.commit()
                status_label.config(text=f"Attendance saved for {date}")
            except Exception as e:
                status_label.config(text=f"Error: {str(e)}")

        def clear_all():
            for emp_num in attendance_vars:
                for var in attendance_vars[emp_num].values():
                    var.set(False)
                notes_entries[emp_num].delete(0, tk.END)
            for var in select_all_vars.values():
                var.set(False)
            status_label.config(text="Form cleared")

        # Layout canvas and scrollbar
        canvas.grid(row=1, column=0, columnspan=6, sticky="nsew")
        scrollbar.grid(row=1, column=6, sticky="ns")
        
        # Buttons and status
        ttk.Button(win, text="Save", command=save_attendance).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(win, text="Clear All", command=clear_all).grid(row=2, column=2, padx=5, pady=5)
        ttk.Button(win, text="Load Existing", command=load_existing).grid(row=2, column=3, padx=5, pady=5)
        status_label = ttk.Label(win, text="")
        status_label.grid(row=3, column=0, columnspan=6, padx=5, pady=5)

        # Configure window resizing
        win.grid_rowconfigure(1, weight=1)
        win.grid_columnconfigure(0, weight=1)

        # Initial load
        load_existing()
        
    def daily_report(self):
            win = tk.Toplevel(self.root)
            win.title("Daily Report")
            win.geometry("600x400")

            date_label = ttk.Label(win, text="Select Date")
            date_label.grid(row=0, column=0, padx=5, pady=5)
            date_entry = ttk.Entry(win)
            date_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
            date_entry.grid(row=0, column=1, padx=5, pady=5)

            tree = ttk.Treeview(win, columns=("Number", "Name", "Dept", "Present", "Tardy", "Early", "Absent", "Bonus", "Notes"), show="headings")
            tree.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

            for col in tree["columns"]:
                tree.heading(col, text=col)
                tree.column(col, width=70)

            def update_report():
             for item in tree.get_children():
                tree.delete(item)
            
            date = date_entry.get()
            c = self.conn.cursor()
            c.execute("""
                SELECT e.employee_number, e.first_name, e.last_name, e.department, 
                       a.present, a.tardy, a.early_out, a.absent, a.bonus_points, a.notes
                FROM employees e
                LEFT JOIN attendance a ON e.employee_number = a.employee_number AND a.date = ?
                WHERE e.termination_date IS NULL""", (date,))
            
            for row in c.fetchall():
                tree.insert("", "end", values=(row[0], f"{row[1]} {row[2]}", row[3], *row[4:]))

            ttk.Button(win, text="Generate Report", command=update_report).grid(row=2, column=1, padx=5, pady=5)
            update_report()  # Initial load

    def weekly_summary(self):
        win = tk.Toplevel(self.root)
        win.title("Weekly Summary")
        win.geometry("600x400")

        start_date = datetime.now() - timedelta(days=datetime.now().weekday())
        ttk.Label(win, text="Week Start Date").grid(row=0, column=0, padx=5, pady=5)
        date_entry = ttk.Entry(win)
        date_entry.insert(0, start_date.strftime("%Y-%m-%d"))
        date_entry.grid(row=0, column=1, padx=5, pady=5)

        tree = ttk.Treeview(win, columns=("Number", "Name", "Dept", "Present", "Tardy", "Early", "Absent", "Bonus"), show="headings")
        tree.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

        for col in tree["columns"]:
            tree.heading(col, text=col)
            tree.column(col, width=70)

        def update_summary():
            for item in tree.get_children():
                tree.delete(item)
            
            start = datetime.strptime(date_entry.get(), "%Y-%m-%d")
            end = start + timedelta(days=6)
            
            c = self.conn.cursor()
            c.execute("""
                SELECT e.employee_number, e.first_name, e.last_name, e.department,
                       SUM(CASE WHEN a.present THEN 1 ELSE 0 END),
                       SUM(CASE WHEN a.tardy THEN 1 ELSE 0 END),
                       SUM(CASE WHEN a.early_out THEN 1 ELSE 0 END),
                       SUM(CASE WHEN a.absent THEN 1 ELSE 0 END),
                       SUM(CASE WHEN a.bonus_points THEN 1 ELSE 0 END)
                FROM employees e
                LEFT JOIN attendance a ON e.employee_number = a.employee_number 
                AND a.date BETWEEN ? AND ?
                WHERE e.termination_date IS NULL
                GROUP BY e.employee_number, e.first_name, e.last_name, e.department""",
                (start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")))
            
            for row in c.fetchall():
                tree.insert("", "end", values=row)

        ttk.Button(win, text="Generate Summary", command=update_summary).grid(row=2, column=1, padx=5, pady=5)
        update_summary()

    def monthly_summary(self):
        win = tk.Toplevel(self.root)
        win.title("Monthly Summary")
        win.geometry("600x400")

        current_month = datetime.now().replace(day=1)
        ttk.Label(win, text="Month Start Date").grid(row=0, column=0, padx=5, pady=5)
        date_entry = ttk.Entry(win)
        date_entry.insert(0, current_month.strftime("%Y-%m-%d"))
        date_entry.grid(row=0, column=1, padx=5, pady=5)

        tree = ttk.Treeview(win, columns=("Number", "Name", "Dept", "Present", "Tardy", "Early", "Absent", "Bonus"), show="headings")
        tree.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

        for col in tree["columns"]:
            tree.heading(col, text=col)
            tree.column(col, width=70)

        def update_summary():
            for item in tree.get_children():
                tree.delete(item)
            
            start = datetime.strptime(date_entry.get(), "%Y-%m-%d")
            end = (start.replace(month=start.month+1, day=1) - timedelta(days=1)) if start.month < 12 else start.replace(year=start.year+1, month=1, day=31)
            
            c = self.conn.cursor()
            c.execute("""
                SELECT e.employee_number, e.first_name, e.last_name, e.department,
                       SUM(CASE WHEN a.present THEN 1 ELSE 0 END),
                       SUM(CASE WHEN a.tardy THEN 1 ELSE 0 END),
                       SUM(CASE WHEN a.early_out THEN 1 ELSE 0 END),
                       SUM(CASE WHEN a.absent THEN 1 ELSE 0 END),
                       SUM(CASE WHEN a.bonus_points THEN 1 ELSE 0 END)
                FROM employees e
                LEFT JOIN attendance a ON e.employee_number = a.employee_number 
                AND a.date BETWEEN ? AND ?
                WHERE e.termination_date IS NULL
                GROUP BY e.employee_number, e.first_name, e.last_name, e.department""",
                (start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")))
            
            for row in c.fetchall():
                tree.insert("", "end", values=row)

        ttk.Button(win, text="Generate Summary", command=update_summary).grid(row=2, column=1, padx=5, pady=5)
        update_summary()

    def import_names(self):
        # Import from CSV
        with open("employees.csv", "r") as file:
            reader = csv.DictReader(file)
            c = self.conn.cursor()
            for row in reader:
                c.execute("INSERT INTO employees (employee_number, first_name, last_name, department, hire_date, termination_date) VALUES (?, ?, ?, ?, ?, ?)",
                (row["employee_number"], row["first_name"], row["last_name"], row["department"], row["hire_date"], row["termination_date"]))
            self.conn.commit()
        messagebox.showinfo("Success", "Names imported!")

    def __del__(self):
        self.conn.close()

# Run the app
if __name__ == "__main__":
    init_db()
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()