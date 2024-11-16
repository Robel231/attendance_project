import sqlite3
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime
from ttkthemes import ThemedTk  # For better themes and styling

# Connect to the database and ensure tables are created
def setup_database():
    conn = sqlite3.connect('attendance.db')
    cursor = conn.cursor()

    # Create Employees table
    cursor.execute('''CREATE TABLE IF NOT EXISTS employees (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        name TEXT NOT NULL,
                        position TEXT,
                        department TEXT,
                        employee_code TEXT UNIQUE NOT NULL
                    )''')

    # Create Attendance table
    cursor.execute('''CREATE TABLE IF NOT EXISTS attendance (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        employee_id INTEGER,
                        clock_in_time TIMESTAMP,
                        clock_out_time TIMESTAMP,
                        date TEXT NOT NULL,
                        FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE
                    )''')

    # Create Users table for login and role management
    cursor.execute('''CREATE TABLE IF NOT EXISTS users (
                        username TEXT PRIMARY KEY,
                        password TEXT NOT NULL,
                        role TEXT NOT NULL CHECK(role IN ('HR', 'Employee'))
                    )''')

    # Insert a default HR user for testing
    cursor.execute("INSERT OR IGNORE INTO users (username, password, role) VALUES (?, ?, ?)",
                   ('admin', 'admin123', 'HR'))

    conn.commit()
    conn.close()

# Add animation for switching screens
def animate_frame(current_frame, target_frame):
    current_frame.pack_forget()  # Hide current frame
    target_frame.pack(fill="both", expand=True)  # Show target frame

# Functions for individual features
def open_employee_registration():
    def save_employee():
        name = name_entry.get()
        position = position_entry.get()
        department = department_entry.get()
        employee_code = code_entry.get()

        if not name or not employee_code:
            messagebox.showwarning("Input Error", "Name and Employee Code are required.")
            return

        try:
            conn = sqlite3.connect('attendance.db')
            cursor = conn.cursor()
            cursor.execute("INSERT INTO employees (name, position, department, employee_code) VALUES (?, ?, ?, ?)",
                           (name, position, department, employee_code))
            conn.commit()
            conn.close()
            messagebox.showinfo("Success", "Employee added successfully.")
            reg_window.destroy()
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", "Employee code must be unique.")
            conn.close()

    reg_window = tk.Toplevel(root)
    reg_window.title("Employee Registration")
    reg_window.geometry("300x300")
    ttk.Label(reg_window, text="Name:").pack(pady=5)
    name_entry = ttk.Entry(reg_window)
    name_entry.pack()
    ttk.Label(reg_window, text="Position:").pack(pady=5)
    position_entry = ttk.Entry(reg_window)
    position_entry.pack()
    ttk.Label(reg_window, text="Department:").pack(pady=5)
    department_entry = ttk.Entry(reg_window)
    department_entry.pack()
    ttk.Label(reg_window, text="Employee Code:").pack(pady=5)
    code_entry = ttk.Entry(reg_window)
    code_entry.pack()
    ttk.Button(reg_window, text="Save Employee", command=save_employee).pack(pady=10)

def open_create_user():
    def save_user():
        username = username_entry.get()
        password = password_entry.get()
        role = role_entry.get().capitalize()

        if not username or not password or role not in ["HR", "Employee"]:
            messagebox.showwarning("Input Error", "All fields are required and role must be 'HR' or 'Employee'.")
            return

        try:
            conn = sqlite3.connect('attendance.db')
            cursor = conn.cursor()
            cursor.execute("INSERT INTO users (username, password, role) VALUES (?, ?, ?)",
                           (username, password, role))
            conn.commit()
            conn.close()
            messagebox.showinfo("Success", "User created successfully.")
            create_user_window.destroy()
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", "Username must be unique.")
            conn.close()

    create_user_window = tk.Toplevel(root)
    create_user_window.title("Create User")
    create_user_window.geometry("300x250")
    ttk.Label(create_user_window, text="Username:").pack(pady=5)
    username_entry = ttk.Entry(create_user_window)
    username_entry.pack()
    ttk.Label(create_user_window, text="Password:").pack(pady=5)
    password_entry = ttk.Entry(create_user_window, show="*")
    password_entry.pack()
    ttk.Label(create_user_window, text="Role (HR or Employee):").pack(pady=5)
    role_entry = ttk.Entry(create_user_window)
    role_entry.pack()
    ttk.Button(create_user_window, text="Save User", command=save_user).pack(pady=10)

def open_delete_employee():
    def delete_employee():
        employee_code = code_entry.get()
        if not employee_code:
            messagebox.showwarning("Input Error", "Employee Code is required.")
            return

        try:
            conn = sqlite3.connect('attendance.db')
            cursor = conn.cursor()
            cursor.execute("SELECT id FROM employees WHERE employee_code = ?", (employee_code,))
            result = cursor.fetchone()
            if result:
                cursor.execute("DELETE FROM employees WHERE id = ?", (result[0],))
                conn.commit()
                messagebox.showinfo("Success", "Employee deleted successfully.")
                delete_window.destroy()
            else:
                messagebox.showerror("Error", "Employee not found.")
            conn.close()
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            conn.close()

    delete_window = tk.Toplevel(root)
    delete_window.title("Delete Employee")
    delete_window.geometry("300x150")
    ttk.Label(delete_window, text="Enter Employee Code to Delete:").pack(pady=5)
    code_entry = ttk.Entry(delete_window)
    code_entry.pack()
    ttk.Button(delete_window, text="Delete Employee", command=delete_employee).pack(pady=10)

def open_attendance():
    attendance_window = tk.Toplevel(root)
    attendance_window.title("Clock In/Out")
    attendance_window.geometry("300x200")

    def clock_in():
        employee_code = code_entry.get()
        if not employee_code:
            messagebox.showwarning("Input Error", "Employee Code is required.")
            return

        conn = sqlite3.connect('attendance.db')
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM employees WHERE employee_code = ?", (employee_code,))
        result = cursor.fetchone()
        if result:
            employee_id = result[0]
            date_today = datetime.now().date()
            cursor.execute("SELECT * FROM attendance WHERE employee_id = ? AND date = ? AND clock_out_time IS NULL",
                           (employee_id, date_today))
            if cursor.fetchone() is None:
                cursor.execute("INSERT INTO attendance (employee_id, clock_in_time, date) VALUES (?, ?, ?)",
                               (employee_id, datetime.now(), date_today))
                conn.commit()
                messagebox.showinfo("Clock In", f"Clocked in at {datetime.now()}")
            else:
                messagebox.showwarning("Error", "Already clocked in for today.")
        else:
            messagebox.showerror("Error", "Employee not found.")
        conn.close()

    def clock_out():
        employee_code = code_entry.get()
        if not employee_code:
            messagebox.showwarning("Input Error", "Employee Code is required.")
            return

        conn = sqlite3.connect('attendance.db')
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM employees WHERE employee_code = ?", (employee_code,))
        result = cursor.fetchone()
        if result:
            employee_id = result[0]
            date_today = datetime.now().date()
            cursor.execute("SELECT id FROM attendance WHERE employee_id = ? AND date = ? AND clock_out_time IS NULL",
                           (employee_id, date_today))
            record = cursor.fetchone()
            if record:
                cursor.execute("UPDATE attendance SET clock_out_time = ? WHERE id = ?",
                               (datetime.now(), record[0]))
                conn.commit()
                messagebox.showinfo("Clock Out", f"Clocked out at {datetime.now()}")
            else:
                messagebox.showwarning("Error", "No clock-in record found for today.")
        else:
            messagebox.showerror("Error", "Employee not found.")
        conn.close()

    ttk.Label(attendance_window, text="Employee Code:").pack(pady=5)
    code_entry = ttk.Entry(attendance_window)
    code_entry.pack()
    ttk.Button(attendance_window, text="Clock In", command=clock_in).pack(pady=5)
    ttk.Button(attendance_window, text="Clock Out", command=clock_out).pack(pady=5)

def open_report():
    report_window = tk.Toplevel(root)
    report_window.title("Attendance Report")
    report_window.geometry("500x500")  # Adjusted size to fit the new elements
    
    # Add filter labels and entry fields
    ttk.Label(report_window, text="Filter by Name:").pack(pady=5)
    name_filter_entry = ttk.Entry(report_window)
    name_filter_entry.pack(pady=5)

    ttk.Label(report_window, text="Filter by Department:").pack(pady=5)
    department_filter_entry = ttk.Entry(report_window)
    department_filter_entry.pack(pady=5)

    ttk.Label(report_window, text="Filter by Date (YYYY-MM-DD):").pack(pady=5)
    date_filter_entry = ttk.Entry(report_window)
    date_filter_entry.pack(pady=5)

    # Add filter by month (using combobox for months)
    ttk.Label(report_window, text="Filter by Month:").pack(pady=5)
    months = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]
    month_filter_combobox = ttk.Combobox(report_window, values=months)
    month_filter_combobox.pack(pady=5)

    # Table for displaying attendance
    report_tree = ttk.Treeview(report_window, columns=("Name", "Department", "Date", "Clock In", "Clock Out"), show="headings")
    report_tree.heading("Name", text="Name")
    report_tree.heading("Department", text="Department")
    report_tree.heading("Date", text="Date")
    report_tree.heading("Clock In", text="Clock In")
    report_tree.heading("Clock Out", text="Clock Out")
    report_tree.pack(fill="both", expand=True)
    
    def search_attendance():
        # Get filter values
        name_filter = name_filter_entry.get().lower()
        department_filter = department_filter_entry.get().lower()
        date_filter = date_filter_entry.get()
        month_filter = month_filter_combobox.get()

        # Connect to the database and apply filters
        conn = sqlite3.connect('attendance.db')
        cursor = conn.cursor()

        query = '''SELECT e.name, e.department, a.date, a.clock_in_time, a.clock_out_time
                   FROM attendance a
                   JOIN employees e ON a.employee_id = e.id
                   WHERE 1=1'''  # Default where clause to apply conditions dynamically

        params = []  # List to hold parameters for the query

        if name_filter:
            query += " AND e.name LIKE ?"
            params.append(f"%{name_filter}%")
        
        if department_filter:
            query += " AND e.department LIKE ?"
            params.append(f"%{department_filter}%")
        
        if date_filter:
            query += " AND a.date = ?"
            params.append(date_filter)
        
        if month_filter:  # Apply filter for month
            month_number = months.index(month_filter) + 1  # Convert month name to number (1-12)
            query += " AND strftime('%m', a.date) = ?"
            params.append(f"{month_number:02d}")  # Ensure two-digit month (e.g., "03" for March)

        query += " ORDER BY a.date DESC"  # Sorting by date

        cursor.execute(query, tuple(params))
        records = cursor.fetchall()
        conn.close()

        # Clear the current data in the treeview
        for row in report_tree.get_children():
            report_tree.delete(row)
        
        # Insert the filtered records into the treeview
        for record in records:
            report_tree.insert("", "end", values=record)

    def export_to_excel():
        # Get data from the Treeview
        data = []
        for row in report_tree.get_children():
            data.append(report_tree.item(row)['values'])
        
        # Convert data into a pandas DataFrame
        df = pd.DataFrame(data, columns=["Name", "Department", "Date", "Clock In", "Clock Out"])
        
        # Define the file path
        file_path = os.path.abspath("attendance_report.xlsx")
        
        # Save to Excel file
        try:
            df.to_excel(file_path, index=False, engine='openpyxl')
            messagebox.showinfo("Success", f"Excel file created successfully!\nPath: {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create Excel file.\nError: {str(e)}")

    # Button to trigger search based on filters
    ttk.Button(report_window, text="Search", command=search_attendance).pack(pady=10)
    
    # Button to export data to Excel
    ttk.Button(report_window, text="Export to Excel", command=export_to_excel).pack(pady=10)


# Main application setup
setup_database()
root = ThemedTk(theme="arc")  # Using a modern theme
root.title("Employee Attendance System")
root.geometry("400x400")

# Login Frame
login_frame = ttk.Frame(root)
login_frame.pack(fill="both", expand=True)
welcome_label = ttk.Label(login_frame, text="Welcome to Attendance System", font=("Arial", 16))
welcome_label.pack(pady=10)

# Login Form
ttk.Label(login_frame, text="Username:").pack(pady=5)
username_entry = ttk.Entry(login_frame)
username_entry.pack()
ttk.Label(login_frame, text="Password:").pack(pady=5)
password_entry = ttk.Entry(login_frame, show="*")
password_entry.pack()

def login():
    username = username_entry.get()
    password = password_entry.get()
    conn = sqlite3.connect('attendance.db')
    cursor = conn.cursor()
    cursor.execute("SELECT role FROM users WHERE username = ? AND password = ?", (username, password))
    result = cursor.fetchone()
    conn.close()
    if result:
        role = result[0]
        if role == "HR":
            animate_frame(login_frame, hr_dashboard_frame)
        elif role == "Employee":
            animate_frame(login_frame, employee_dashboard_frame)
    else:
        messagebox.showerror("Login Error", "Invalid username or password.")

ttk.Button(login_frame, text="Login", command=login).pack(pady=20)

# HR Dashboard
hr_dashboard_frame = ttk.Frame(root)
ttk.Label(hr_dashboard_frame, text="HR Dashboard", font=("Arial", 16)).pack(pady=10)
ttk.Button(hr_dashboard_frame, text="Employee Registration", command=open_employee_registration, width=25).pack(pady=10)
ttk.Button(hr_dashboard_frame, text="Clock In/Out", command=open_attendance, width=25).pack(pady=10)
ttk.Button(hr_dashboard_frame, text="Attendance Report", command=open_report, width=25).pack(pady=10)
ttk.Button(hr_dashboard_frame, text="Create User", command=open_create_user, width=25).pack(pady=10)
ttk.Button(hr_dashboard_frame, text="Delete Employee", command=open_delete_employee, width=25).pack(pady=10)
ttk.Button(hr_dashboard_frame, text="Logout", command=lambda: animate_frame(hr_dashboard_frame, login_frame), width=25).pack(pady=10)

# Employee Dashboard
employee_dashboard_frame = ttk.Frame(root)
ttk.Label(employee_dashboard_frame, text="Employee Dashboard", font=("Arial", 16)).pack(pady=10)
ttk.Button(employee_dashboard_frame, text="Clock In/Out", command=open_attendance, width=25).pack(pady=10)
#ttk.Button(employee_dashboard_frame, text="Attendance Report", command=open_report, width=25).pack(pady=10)
ttk.Button(employee_dashboard_frame, text="Logout", command=lambda: animate_frame(employee_dashboard_frame, login_frame), width=25).pack(pady=10)

root.mainloop()
