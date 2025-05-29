import uuid
import pandas as pd
import os
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

class Department:
    def __init__(self, name, employees=None):
        if employees is None:
            self.employees = dict()
        else:
            self.employees = employees
        self.name       = name

    def add_employee(self, employee):
        self.employees[employee.uuid] = employee

    def remove_employee(self, employee):
        self.employees.pop(employee.uuid)

class Employee:
    def __init__(self, name):
        self.uuid = uuid.uuid4()
        self.name = name

def load_department_from_file(filename):
    department      = None
    department_list = []
    with open(filename) as f:
        for line in f:
            line_str = line.strip()
            if line_str != '':
                line_str = line_str.rstrip()
                if '•' in line_str:
                    employee_name   = line_str.split('•')[1].replace('\t','')
                    employee        = Employee(employee_name)
                    try:
                        department.add_employee(employee)
                    except Exception as e:
                        print(e)
                else:
                    if department is not None:
                        department_list.append(department)
                    department = Department(line_str)

    return department_list


departments = load_department_from_file('README.txt')
dept_employee_map = dict()
for department in departments:
    dept_employee_map[department.name] = [emp.name for emp in department.employees.values()]



# ==== STEP 1: Load Inventory ====
inventory_path = "equipment-Inventory.xlsx"
log_path = "device_borrow_log.xlsx"

if not os.path.exists(inventory_path):
    messagebox.showerror("Missing File", "equipment-Inventory.xlsx not found.")
    raise FileNotFoundError("equipment-Inventory.xlsx not found.")

inventory_df = pd.read_excel(inventory_path)
inventory_df = inventory_df.fillna("N/A")

device_dict = {}
for _, row in inventory_df.iterrows():
    key = f"{row['DESCRIPTION']} ({row['EQUIPMENT TYPE']})"
    device_dict[key] = {
        "model": str(row["ID CODE/ MODEL NO."]),
        "vendor": row["VENDOR"],
        "location": row["STOCK LOCATION"],
        "status": row["STATUS"],
        "last_cleaned": str(row["LAST CLEAN DATE"]),
        "number_in_stock": row["NUMBER IN STOCK"]
    }

# ==== GUI Setup ====
root = tk.Tk()
root.title("Equipment Borrowing System")
root.geometry("650x600")
FONT = ("Segoe UI", 11)

device_name_var = tk.StringVar()
model_var = tk.StringVar()
vendor_var = tk.StringVar()
location_var = tk.StringVar()
status_var = tk.StringVar()
last_cleaned_var = tk.StringVar()

department_selection_var = tk.StringVar()
employee_selection_var = tk.StringVar()


name_var = tk.StringVar()
dept_var = tk.StringVar()
employee_search_var = tk.StringVar()

# ==== Autofill on device select ====
def update_device_info(*args):
    selected = device_name_var.get()
    data = device_dict.get(selected, {})
    model_var.set(data.get("model", "N/A"))
    vendor_var.set(data.get("vendor", "N/A"))
    location_var.set(data.get("location", "N/A"))
    status_var.set(data.get("status", "N/A"))
    last_cleaned_var.set(data.get("last_cleaned", "N/A"))

# ==== GUI Components ====
def add_label_entry(parent, row, label, var, readonly=False):
    tk.Label(parent, text=label, font=FONT).grid(row=row, column=0, sticky="e", padx=10, pady=5)
    state = "readonly" if readonly else "normal"
    entry = tk.Entry(parent, textvariable=var, font=FONT, width=35, state=state)
    entry.grid(row=row, column=1, padx=10, pady=5)

# ==== Borrow / Return Logic ====
def log_action(action):
    name = name_var.get().strip()
    dept = dept_var.get().strip()
    device = device_name_var.get()
    model = model_var.get()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if not name or not dept:
        messagebox.showwarning("Missing Info", "Please enter your name and department.")
        return

    # Load or create log
    columns = ["Name", "Department", "Device Name", "Model", "Timestamp", "Action"]
    if os.path.exists(log_path):
        log_df = pd.read_excel(log_path)
    else:
        log_df = pd.DataFrame(columns=columns)

    # If RETURN: check that this user has borrowed it
    if action == "Return":
        user_log = log_df[(log_df["Name"] == name) & (log_df["Device Name"] == device)]
        if user_log.empty or (user_log["Action"] == "Borrow").sum() <= (user_log["Action"] == "Return").sum():
            messagebox.showwarning("Invalid Return", f"No active borrow record for {device} by {name}.")
            return

    # If BORROW: check device not already borrowed
    if action == "Borrow":
        device_log = log_df[log_df["Device Name"] == device]
        if (device_log["Action"] == "Borrow").sum() > (device_log["Action"] == "Return").sum():
            messagebox.showwarning("Device In Use", f"{device} is currently borrowed by someone else.")
            return

    # Append log
    new_row = pd.DataFrame([[name, dept, device, model, timestamp, action]], columns=columns)
    log_df = pd.concat([log_df, new_row], ignore_index=True)
    log_df.to_excel(log_path, index=False)
    messagebox.showinfo("Success", f"{action} recorded for {device} by {name}.")
    name_var.set("")
    dept_var.set("")
    device_name_var.set(list(device_dict.keys())[0])
    update_device_info()


# ==== View Employee Borrowed Devices ====
def show_borrowed_equipment():
    emp_name = employee_search_var.get().strip()
    if not emp_name:
        messagebox.showwarning("Missing Name", "Please select an employee.")
        return

    if not os.path.exists(log_path):
        messagebox.showinfo("No Data", "No logs found yet.")
        return

    log_df = pd.read_excel(log_path)
    user_log = log_df[log_df["Name"] == emp_name]

    borrowed = []
    for device in user_log["Device Name"].unique():
        d_log = user_log[user_log["Device Name"] == device]
        if (d_log["Action"] == "Borrow").sum() > (d_log["Action"] == "Return").sum():
            borrowed.append(device)

    result = "\n".join(borrowed) if borrowed else "No borrowed equipment found."
    popup = tk.Toplevel(root)
    popup.title(f"Equipment borrowed by {emp_name}")
    tk.Label(popup, text=result, font=FONT, justify="left", wraplength=500).pack(padx=20, pady=20)

# ==== GUI Layout ====
add_label_entry(root, 0, "Employee Name:", name_var)
add_label_entry(root, 1, "Department:", dept_var)

tk.Label(root, text="Device Name:", font=FONT).grid(row=2, column=0, sticky="e", padx=10, pady=5)
device_dropdown = ttk.OptionMenu(root, device_name_var, list(device_dict.keys())[0], *device_dict.keys())
device_dropdown.grid(row=2, column=1, padx=10, pady=5, sticky="w")

add_label_entry(root, 3, "Model:", model_var, True)
add_label_entry(root, 4, "Vendor:", vendor_var, True)
add_label_entry(root, 5, "Location:", location_var, True)
add_label_entry(root, 6, "Status:", status_var, True)
add_label_entry(root, 7, "Last Cleaned:", last_cleaned_var, True)

# Buttons
tk.Button(root, text="Borrow", command=lambda: log_action("Borrow"),
          bg="#4CAF50", fg="white", font=FONT, width=20).grid(row=8, column=0, columnspan=2, pady=10)

tk.Button(root, text="Return", command=lambda: log_action("Return"),
          bg="#2196F3", fg="white", font=FONT, width=20).grid(row=9, column=0, columnspan=2, pady=5)

# ==== Employee Viewer ====
tk.Label(root, text="Select Employee:", font=FONT).grid(row=10, column=0, sticky="e", padx=10, pady=5)
if os.path.exists(log_path):
    try:
        log_df = pd.read_excel(log_path)
        employees = sorted(log_df["Name"].dropna().unique())
    except:
        employees = []
else:
    employees = []

employee_dropdown = ttk.OptionMenu(root, employee_search_var, employees[0] if employees else "", *employees)
employee_dropdown.grid(row=10, column=1, padx=10, pady=5, sticky="w")

tk.Button(root, text="View Borrowed Equipment", command=show_borrowed_equipment,
          bg="#FF9800", fg="white", font=FONT, width=25).grid(row=11, column=0, columnspan=2, pady=20)

# Setup auto-update for metadata
device_name_var.trace_add("write", update_device_info)
update_device_info()

## DEPT-EMPLOYEE DROPDOWN

def update_employee_dropdown(*args):
    selected_dept   = department_selection_var.get()
    employees       = dept_employee_map.get(selected_dept, [])
    menu            = employee_dropdown['menu']
    menu.delete(0, 'end')

    if employees:
        for emp in employees:
            menu.add_command(label=emp, command=lambda val=emp: employee_selection_var.set(val))
        employee_selection_var.set(employees[0])
    else:
        employee_selection_var.set("")

# Department Dropdown
tk.Label(root, text="Department:", font=FONT).grid(row=12, column=0, sticky="e", padx=10, pady=5)
dept_options = list(dept_employee_map.keys())
department_selection_var.set(dept_options[0] if dept_options else "")
department_dropdown = ttk.OptionMenu(root, department_selection_var, department_selection_var.get(), *dept_options)
department_dropdown.grid(row=12, column=1, padx=10, pady=5, sticky="w")

# Employee Dropdown (updates when department changes)
tk.Label(root, text="Employee:", font=FONT).grid(row=13, column=0, sticky="e", padx=10, pady=5)
employee_dropdown = ttk.OptionMenu(root, employee_selection_var, "")
employee_dropdown.grid(row=13, column=1, padx=10, pady=5, sticky="w")

# Trigger update
department_selection_var.trace_add("write", update_employee_dropdown)
update_employee_dropdown()

selected_employee = employee_selection_var.get()


root.mainloop()
