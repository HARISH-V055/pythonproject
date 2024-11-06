import sqlite3
import tkinter as tk
import csv
import os

conn = sqlite3.connect('employee_management.db')
c = conn.cursor()

c.execute('''CREATE TABLE IF NOT EXISTS employees (
        emp_id INTEGER PRIMARY KEY,
        name TEXT NOT NULL,
        department TEXT,
        role TEXT,
        salary REAL
    )
''')
conn.commit()

def add_employee(emp_id, name, department, role, salary):
    try:
        c.execute("INSERT INTO employees (emp_id, name, department, role, salary) VALUES (?, ?, ?, ?, ?)",
                  (emp_id, name, department, role, salary))
        conn.commit()
        display_message(f"Employee {name} added successfully.")
    except sqlite3.IntegrityError:
        display_message("Employee ID already exists or invalid data.")

def remove_employee_by_id(emp_id):
    c.execute("DELETE FROM employees WHERE emp_id = ?", (emp_id,))
    conn.commit()
    display_message(f"Employee ID {emp_id} removed successfully.")

def remove_employee_by_name(name):
    c.execute("DELETE FROM employees WHERE name = ?", (name,))
    conn.commit()
    display_message(f"Employee(s) with name '{name}' removed successfully.")

def search_employee_by_id(emp_id):
    c.execute("SELECT * FROM employees WHERE emp_id = ?", (emp_id,))
    result = c.fetchone()
    if result:
        display_message(f"ID: {result[0]}, Name: {result[1]}, Dept: {result[2]}, Role: {result[3]}, Salary: ${result[4]}")
    else:
        display_message(f"No employee found with ID {emp_id}.")

def search_employee_by_name(name):
    c.execute("SELECT * FROM employees WHERE name LIKE ?", ('%' + name + '%',))
    results = c.fetchall()
    if results:
        display_message("\n".join(f"ID: {r[0]}, Name: {r[1]}, Dept: {r[2]}, Role: {r[3]}, Salary: ${r[4]}" for r in results))
    else:
        display_message(f"No employee found with name containing '{name}'.")

def display_all_employees():
    c.execute("SELECT * FROM employees")
    results = c.fetchall()
    if results:
        display_message("\n".join(f"ID: {r[0]}, Name: {r[1]}, Dept: {r[2]}, Role: {r[3]}, Salary: ${r[4]}" for r in results))
    else:
        display_message("No employees in the database.")

def update_employee_salary_by_id(emp_id, new_salary):
    c.execute("UPDATE employees SET salary = ? WHERE emp_id = ?", (new_salary, emp_id))
    conn.commit()
    display_message(f"Updated salary of Employee ID {emp_id} to ${new_salary}")

def update_employee_salary_by_name(name, new_salary):
    c.execute("UPDATE employees SET salary = ? WHERE name LIKE ?", (new_salary, '%' + name + '%'))
    conn.commit()
    display_message(f"Updated salary of Employee(s) with name '{name}' to ${new_salary}")

def export_to_csv():
    c.execute("SELECT * FROM employees")
    results = c.fetchall()
    
    with open('employee_data.csv', 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["Employee ID", "Name", "Department", "Role", "Salary"])
        writer.writerows(results)
    
    try:
        os.startfile('employee_data.csv') 
        display_message("Employee data exported to 'employee_data.csv' and opened in Excel.")
    except Exception as e:
        display_message(f"Failed to open in Excel. Error: {str(e)}")

def display_message(message):
    display_area.config(state='normal')
    display_area.delete(1.0, tk.END)
    display_area.insert(tk.END, message)
    display_area.config(state='disabled')

def reset_fields(*fields):
    for field in fields:
        field.delete(0, tk.END)

def admin_menu():
    admin_window = tk.Toplevel()
    admin_window.title("Employee Management System")
    admin_window.geometry("800x400")

    left_frame = tk.Frame(admin_window)
    left_frame.pack(side="left", fill="y", padx=10, pady=10)

    right_frame = tk.Frame(admin_window)
    right_frame.pack(side="right", expand=True, fill="both", padx=10, pady=10)

    global display_area
    display_area = tk.Text(right_frame, wrap="word", state="disabled")
    display_area.pack(expand=True, fill="both")

    emp_id_entry = tk.Entry(left_frame)
    name_entry = tk.Entry(left_frame)
    department_entry = tk.Entry(left_frame)
    role_entry = tk.Entry(left_frame)
    salary_entry = tk.Entry(left_frame)
    search_id_entry = tk.Entry(left_frame)
    search_name_entry = tk.Entry(left_frame)
    update_salary_entry = tk.Entry(left_frame)
    update_salary_by_name_entry = tk.Entry(left_frame)
    update_salary_by_id_entry = tk.Entry(left_frame)

    def add_employee_cmd():
        emp_id = int(emp_id_entry.get())
        name = name_entry.get()
        department = department_entry.get()
        role = role_entry.get()
        salary = float(salary_entry.get())
        add_employee(emp_id, name, department, role, salary)
        reset_fields(emp_id_entry, name_entry, department_entry, role_entry, salary_entry)

    def remove_employee_by_id_cmd():
        emp_id = int(search_id_entry.get())
        remove_employee_by_id(emp_id)
        reset_fields(search_id_entry)

    def remove_employee_by_name_cmd():
        name = search_name_entry.get()
        remove_employee_by_name(name)
        reset_fields(search_name_entry)

    def search_by_id_cmd():
        emp_id = int(search_id_entry.get())
        search_employee_by_id(emp_id)

    def search_by_name_cmd():
        name = search_name_entry.get()
        search_employee_by_name(name)

    def display_all_cmd():
        display_all_employees()

    def update_salary_by_id_cmd():
        emp_id = int(update_salary_by_id_entry.get())
        new_salary = float(update_salary_entry.get())
        update_employee_salary_by_id(emp_id, new_salary)
        reset_fields(update_salary_by_id_entry, update_salary_entry)

    def update_salary_by_name_cmd():
        name = update_salary_by_name_entry.get()
        new_salary = float(update_salary_entry.get())
        update_employee_salary_by_name(name, new_salary)
        reset_fields(update_salary_by_name_entry, update_salary_entry)

    tk.Label(left_frame, text="Employee ID:").pack()
    emp_id_entry.pack()
    tk.Label(left_frame, text="Name:").pack()
    name_entry.pack()
    tk.Label(left_frame, text="Department:").pack()
    department_entry.pack()
    tk.Label(left_frame, text="Role:").pack()
    role_entry.pack()
    tk.Label(left_frame, text="Salary:").pack()
    salary_entry.pack()
    tk.Button(left_frame, text="Add Employee", command=add_employee_cmd).pack(pady=5)

    tk.Label(left_frame, text="Search/Remove by ID:").pack()
    search_id_entry.pack()
    tk.Button(left_frame, text="Remove by ID", command=remove_employee_by_id_cmd).pack(pady=5)
    tk.Button(left_frame, text="Search by ID", command=search_by_id_cmd).pack(pady=5)

    tk.Label(left_frame, text="Search/Remove by Name:").pack()
    search_name_entry.pack()
    tk.Button(left_frame, text="Remove by Name", command=remove_employee_by_name_cmd).pack(pady=5)
    tk.Button(left_frame, text="Search by Name", command=search_by_name_cmd).pack(pady=5)

    tk.Label(left_frame, text="Update Salary by ID:").pack()
    update_salary_by_id_entry.pack()
    tk.Label(left_frame, text="New Salary:").pack()
    update_salary_entry.pack()
    tk.Button(left_frame, text="Update Salary by ID", command=update_salary_by_id_cmd).pack(pady=5)

    tk.Label(left_frame, text="Update Salary by Name:").pack()
    update_salary_by_name_entry.pack()
    tk.Label(left_frame, text="New Salary:").pack()
    update_salary_entry.pack()
    tk.Button(left_frame, text="Update Salary by Name", command=update_salary_by_name_cmd).pack(pady=5)

    tk.Button(left_frame, text="Display All Employees", command=display_all_cmd).pack(pady=10)

    tk.Button(left_frame, text="Detailed View", command=export_to_csv).pack(pady=10)

def main_menu():
    main_window = tk.Tk()
    main_window.title("Employee Management System")
    main_window.geometry("400x200")

    def admin_login_cmd():
        admin_menu()

    tk.Label(main_window, text="Welcome to Employee Management System").pack(pady=10)
    tk.Button(main_window, text="Enter Admin Menu", command=admin_login_cmd).pack(pady=20)

    main_window.mainloop()

if __name__ == "__main__":
    main_menu()