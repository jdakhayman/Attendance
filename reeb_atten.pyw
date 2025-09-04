"""
Employee Attendance System for Reeb Exterior.

This module provides a GUI application to manage employee records and attendance,
with functionality to generate daily reports, and export them to CSV or copy as tables for email.
"""
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import csv
import pyperclip
import os

def init_db():
    """Initialize the SQLite database with employees and attendance tables."""
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

class AttendanceApp:
    """Main application class for managing employee attendance."""
    def __init__(self, tk_root):
        self.root = tk_root
        self.root.title("Reeb Exterior Employee Attendance System")
        self.conn = sqlite3.connect("attendance.db")
        
        self.main_frame = ttk.Frame(tk_root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        ttk.Button(self.main_frame, text="Add Employee", command=self.add_employee).grid(row=0, column=0, pady=5)
        ttk.Button(self.main_frame, text="Edit Employee", command=self.edit_employee).grid(row=0, column=1, pady=5)
        ttk.Button(self.main_frame, text="Record Attendance", command=self.record_attendance).grid(row=0, column=2, pady=5)
        ttk.Button(self.main_frame, text="Daily Report", command=self.daily_report).grid(row=0, column=3, pady=5)
        ttk.Button(self.main_frame, text="Import Names", command=self.import_names).grid(row=0, column=4, pady=5)

    def validate_date(self, date_str):
        """Validate that a date string is in YYYY-MM-DD format and is a valid date."""
        if not date_str:  # Allow empty string for optional fields
            return True
        try:
            datetime.strptime(date_str, "%Y-%m-%d")
            return True
        except ValueError:
            return False

    def add_employee(self):
        """Open a window to add a new employee to the database."""
        win = tk.Toplevel(self.root)
        win.title("Add Employee")
        win.geometry("300x250")
    
        ttk.Label(win, text="Employee Number").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        emp_num = ttk.Entry(win)
        emp_num.grid(row=0, column=1, padx=5, pady=5)
        emp_num.focus()
        
        ttk.Label(win, text="First Name").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        fname = ttk.Entry(win)
        fname.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(win, text="Last Name").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        lname = ttk.Entry(win)
        lname.grid(row=2, column=1, padx=5, pady=5)
        
        ttk.Label(win, text="Department").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        departments = ["PREP", "ACRO", "RKG", "SDL", "EPL"]
        dept = ttk.Combobox(win, values=departments, state="readonly")
        dept.grid(row=3, column=1, padx=5, pady=5)

        ttk.Label(win, text="Hire Date").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
        hdate = ttk.Entry(win)
        hdate.grid(row=4, column=1, padx=5, pady=5)
        
        ttk.Label(win, text="Termination Date (optional)").grid(row=5, column=0, padx=5, pady=5, sticky=tk.W)
        tdate = ttk.Entry(win)
        tdate.grid(row=5, column=1, padx=5, pady=5)
        
        def save_employee():
            emp_num_val = emp_num.get().strip()
            fname_val = fname.get().strip()
            lname_val = lname.get().strip()
            dept_val = dept.get()
            hdate_val = hdate.get().strip()
            tdate_val = tdate.get().strip() or None

            # Required fields validation
            if not emp_num_val or not fname_val or not lname_val or not dept_val or not hdate_val:
                messagebox.showerror("Error", "Required fields (except Termination Date) must be filled!")
                return
            
            # Employee number validation
            try:
                emp_num_int = int(emp_num_val)
            except ValueError:
                messagebox.showerror("Error", "Employee Number must be an integer!")
                return
            
            # Date validation
            if not self.validate_date(hdate_val):
                messagebox.showerror("Error", "Hire Date must be in YYYY-MM-DD format!")
                return
            if tdate_val and not self.validate_date(tdate_val):
                messagebox.showerror("Error", "Termination Date must be in YYYY-MM-DD format!")
                return
            
            try:
                c = self.conn.cursor()
                c.execute("""
                    INSERT INTO employees (employee_number, first_name, last_name, department, hire_date, termination_date)
                    VALUES (?, ?, ?, ?, ?, ?)""",
                    (emp_num_int, fname_val, lname_val, dept_val, hdate_val, tdate_val))
                self.conn.commit()
                messagebox.showinfo("Success", f"Employee {fname_val} {lname_val} (#{emp_num_int}) added!")
                emp_num.delete(0, tk.END)
                fname.delete(0, tk.END)
                lname.delete(0, tk.END)
                dept.set("")
                hdate.delete(0, tk.END)
                tdate.delete(0, tk.END)
                emp_num.focus()
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", f"Employee Number {emp_num_val} already exists!")
            except sqlite3.Error as e:
                messagebox.showerror("Error", f"Database error: {str(e)}")
        
        ttk.Button(win, text="Save", command=save_employee).grid(row=6, column=1, padx=5, pady=10)
        ttk.Button(win, text="Cancel", command=win.destroy).grid(row=6, column=0, padx=5, pady=10)

    def edit_employee(self):
        """Open a window to edit an employee's termination date."""
        win = tk.Toplevel(self.root)
        win.title("Edit Employee Termination Date")
        win.geometry("400x200")

        ttk.Label(win, text="Employee").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        c = self.conn.cursor()
        c.execute("SELECT employee_number, first_name, last_name, department, termination_date FROM employees ORDER BY last_name, first_name")
        employees = [(f"{row[1]} {row[2]} (#{row[0]}) - {row[3]}{' - Terminated: ' + row[4] if row[4] else ''}", row[0]) for row in c.fetchall()]
        emp_var = tk.StringVar()
        emp_dropdown = ttk.Combobox(win, textvariable=emp_var, values=[e[0] for e in employees], state="readonly", width=40)
        emp_dropdown.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(win, text="Termination Date").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        tdate = ttk.Entry(win)
        tdate.grid(row=1, column=1, padx=5, pady=5)

        def load_current():
            if emp_dropdown.get():
                emp_num = employees[emp_dropdown.current()][1]
                c.execute("SELECT termination_date FROM employees WHERE employee_number = ?", (emp_num,))
                current_date = c.fetchone()[0]
                tdate.delete(0, tk.END)
                tdate.insert(0, current_date or "")

        emp_dropdown.bind("<<ComboboxSelected>>", lambda e: load_current())

        def save_termination():
            if not emp_dropdown.get():
                messagebox.showerror("Error", "Please select an employee!")
                return
            
            emp_num = employees[emp_dropdown.current()][1]
            tdate_val = tdate.get().strip() or None
            
            # Date validation
            if tdate_val and not self.validate_date(tdate_val):
                messagebox.showerror("Error", "Termination Date must be in YYYY-MM-DD format!")
                return
            
            try:
                c = self.conn.cursor()
                c.execute("UPDATE employees SET termination_date = ? WHERE employee_number = ?", (tdate_val, emp_num))
                self.conn.commit()
                messagebox.showinfo("Success", f"Termination date updated for employee #{emp_num}")
                load_current()
                c.execute("SELECT employee_number, first_name, last_name, department, termination_date FROM employees ORDER BY last_name, first_name")
                employees[:] = [(f"{row[1]} {row[2]} (#{row[0]}) - {row[3]}{' - Terminated: ' + row[4] if row[4] else ''}", row[0]) for row in c.fetchall()]
                emp_dropdown['values'] = [e[0] for e in employees]
            except sqlite3.Error as e:
                messagebox.showerror("Error", f"Database error: {str(e)}")

        ttk.Button(win, text="Save", command=save_termination).grid(row=2, column=1, padx=5, pady=10)
        ttk.Button(win, text="Cancel", command=win.destroy).grid(row=2, column=0, padx=5, pady=10)

    def record_attendance(self):
        """Open a window to record attendance for employees."""
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
            ORDER BY department, last_name
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
        ttk.Checkbutton(scrollable_frame, text="All Point", variable=select_all_vars['bonus'],
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
            for _, vars_dict in attendance_vars.items():
                vars_dict[category].set(state)

        def load_existing():
            date = date_entry.get()
            c.execute("SELECT employee_number, present, tardy, early_out, absent, bonus_points, notes FROM attendance WHERE date = ?", (date,))
            records = {row[0]: row[1:] for row in c.fetchall()}
            
            for emp_num, vars_dict in attendance_vars.items():
                if emp_num in records:
                    record = records[emp_num]
                    vars_dict['present'].set(record[0])
                    vars_dict['tardy'].set(record[1])
                    vars_dict['early'].set(record[2])
                    vars_dict['absent'].set(record[3])
                    vars_dict['bonus'].set(record[4])
                    notes_entries[emp_num].delete(0, tk.END)
                    notes_entries[emp_num].insert(0, record[5] or "")
                else:
                    # Clear if no record exists
                    for var in vars_dict.values():
                        var.set(False)
                    notes_entries[emp_num].delete(0, tk.END)
            
            # Update select all checkboxes based on current state
            for category, var in select_all_vars.items():
                all_checked = all(vars_dict[category].get() for vars_dict in attendance_vars.values())
                var.set(all_checked)

        def save_attendance():
            date = date_entry.get()
            if not self.validate_date(date):
                messagebox.showerror("Error", "Date must be in YYYY-MM-DD format!")
                return
                
            c = self.conn.cursor()
            try:
                for emp_num, vars_dict in attendance_vars.items():
                    c.execute("""
                        INSERT OR REPLACE INTO attendance 
                        (employee_number, date, present, tardy, early_out, absent, bonus_points, notes)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                        (emp_num, date,
                         vars_dict['present'].get(),
                         vars_dict['tardy'].get(),
                         vars_dict['early'].get(),
                         vars_dict['absent'].get(),
                         vars_dict['bonus'].get(),
                         notes_entries[emp_num].get()))
                self.conn.commit()
                status_label.config(text=f"Attendance saved for {date}")
            except sqlite3.Error as e:
                status_label.config(text=f"Database error: {str(e)}")

        def clear_all():
            for emp_num, vars_dict in attendance_vars.items():
                for var in vars_dict.values():
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

    def export_to_csv(self, data, headers, filename):
        """Export report data to a CSV file in the user's home directory."""
        try:
            home_dir = os.path.expanduser("~")
            file_path = os.path.join(home_dir, filename)
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                for row in data:
                    writer.writerow(row)
            messagebox.showinfo("Success", f"Report exported to {file_path}")
        except IOError as e:
            messagebox.showerror("Error", f"Failed to export CSV: {str(e)}")

    def copy_to_clipboard(self, data, headers):
        """Copy report data as a fixed-width table for email."""
        # Define column widths for alignment
        col_widths = {
            "Number": 8,
            "Name": 22,
            "Dept": 8,
            "Present": 8,
            "Tardy": 7,
            "Early": 7,
            "Absent": 7,
            "Point": 7,
            "Notes": 30
        }
        
        # Format headers with extra padding
        table = "".join(f"{header:<{col_widths[header]}} " for header in headers).rstrip() + "\n"
        table += "".join("-" * col_widths[header] + " " for header in headers).rstrip() + "\n"
        
        # Format data rows with truncation
        for row in data:
            table += "".join(f"{str(x)[:col_widths[headers[i]]-1]:<{col_widths[headers[i]]}} " for i, x in enumerate(row)).rstrip() + "\n"
        
        pyperclip.copy(table)
        messagebox.showinfo("Success", "Table copied to clipboard!")

    def daily_report(self):
        """Generate and display a daily attendance report."""
        win = tk.Toplevel(self.root)
        win.title("Daily Report")
        win.geometry("900x600")

        date_label = ttk.Label(win, text="Select Date")
        date_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        date_entry = ttk.Entry(win)
        date_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
        date_entry.grid(row=0, column=1, padx=5, pady=5)

        tree = ttk.Treeview(win, columns=("Number", "Name", "Dept", "Present", "Tardy", "Early", "Absent", "Point", "Notes"), show="headings")
        tree.grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky="nsew")

        # Set column widths to match copied table
        col_widths = {"Number": 80, "Name": 176, "Dept": 80, "Present": 70, "Tardy": 60, "Early": 60, "Absent": 60, "Point": 60, "Notes": 240}
        for col in tree["columns"]:
            tree.heading(col, text=col)
            tree.column(col, width=col_widths[col], anchor="w")

        # Scrollbar
        scrollbar = ttk.Scrollbar(win, orient="vertical", command=tree.yview)
        scrollbar.grid(row=1, column=4, sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)

        def update_report():
            for item in tree.get_children():
                tree.delete(item)
            
            date = date_entry.get()
            if not self.validate_date(date):
                messagebox.showerror("Error", "Date must be in YYYY-MM-DD format!")
                return

            c = self.conn.cursor()
            c.execute("""
                SELECT e.employee_number, e.first_name, e.last_name, e.department, 
                       a.present, a.tardy, a.early_out, a.absent, a.bonus_points, a.notes
                FROM employees e
                LEFT JOIN attendance a ON e.employee_number = a.employee_number AND a.date = ?
                WHERE e.termination_date IS NULL
                ORDER BY e.department, e.last_name""", (date,))
            
            data = []
            for row in c.fetchall():
                display_row = (
                    row[0], 
                    f"{row[1]} {row[2]}", 
                    row[3],
                    "Yes" if row[4] else "No",
                    "Yes" if row[5] else "No",
                    "Yes" if row[6] else "No",
                    "Yes" if row[7] else "No",
                    "Yes" if row[8] else "No",
                    row[9] or ""
                )
                tree.insert("", "end", values=display_row)
                data.append(display_row)

            # Update export buttons
            export_csv_btn.config(command=lambda: self.export_to_csv(
                data, tree["columns"], f"daily_report_{date}.csv"))
            copy_btn.config(command=lambda: self.copy_to_clipboard(data, tree["columns"]))

        # Buttons
        ttk.Button(win, text="Generate Report", command=update_report).grid(row=2, column=0, padx=5, pady=5)
        export_csv_btn = ttk.Button(win, text="Export to CSV", command=lambda: None)
        export_csv_btn.grid(row=2, column=1, padx=5, pady=5)
        copy_btn = ttk.Button(win, text="Copy Table", command=lambda: None)
        copy_btn.grid(row=2, column=2, padx=5, pady=5)
        ttk.Button(win, text="Close", command=win.destroy).grid(row=2, column=3, padx=5, pady=5)

        # Configure window resizing
        win.grid_rowconfigure(1, weight=1)
        win.grid_columnconfigure(0, weight=1)

        update_report()

    def import_names(self):
        """Import employee data from a CSV file."""
        try:
            with open("employees.csv", "r", encoding='utf-8') as file:
                reader = csv.DictReader(file)
                c = self.conn.cursor()
                for row in reader:
                    c.execute("""
                        INSERT OR IGNORE INTO employees 
                        (employee_number, first_name, last_name, department, hire_date, termination_date) 
                        VALUES (?, ?, ?, ?, ?, ?)""",
                        (row["employee_number"], row["first_name"], row["last_name"], 
                         row["department"], row["hire_date"], row.get("termination_date")))
                self.conn.commit()
                messagebox.showinfo("Success", "Names imported successfully!")
        except (IOError, sqlite3.Error) as e:
            messagebox.showerror("Error", f"Failed to import names: {str(e)}")

    def __del__(self):
        self.conn.close()

if __name__ == "__main__":
    init_db()
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()