import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
import os
from equipment_library import load_inventory, get_department_employee_map

# File paths
inventory_path = "equipment-Inventory.xlsx"
log_path = "device_borrow_log.xlsx"

# Load inventory and departments
device_dict = load_inventory(inventory_path)
dept_employee_map = get_department_employee_map()

# GUI setup
root = tk.Tk()
root.title("Equipment Borrowing System")
root.geometry("700x700")
FONT = ("Segoe UI", 11)

# Variables for borrow/return section
dept_var = tk.StringVar()
emp_var = tk.StringVar()
device_name_var = tk.StringVar()
model_var = tk.StringVar()
vendor_var = tk.StringVar()
location_var = tk.StringVar()
status_var = tk.StringVar()
last_cleaned_var = tk.StringVar()

# Variables for view borrowed equipment section
view_dept_var = tk.StringVar()
view_emp_var = tk.StringVar()

# ---- Helper Functions ----
def update_device_info(*args):
    selected = device_name_var.get()
    data = device_dict.get(selected, {})
    model_var.set(data.get("model", "N/A"))
    vendor_var.set(data.get("vendor", "N/A"))
    location_var.set(data.get("location", "N/A"))
    status_var.set(data.get("status", "N/A"))
    last_cleaned_var.set(data.get("last_cleaned", "N/A"))

def update_employee_dropdown(dept_var, emp_var, dropdown):
    selected_dept = dept_var.get()
    employees = dept_employee_map.get(selected_dept, [])
    menu = dropdown["menu"]
    menu.delete(0, "end")
    for emp in employees:
        menu.add_command(label=emp, command=lambda val=emp: emp_var.set(val))
    if employees:
        emp_var.set(employees[0])
    else:
        emp_var.set("")

def log_action(action):
    name = emp_var.get().strip()
    dept = dept_var.get().strip()
    device = device_name_var.get()
    model = model_var.get()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if not name or not dept:
        messagebox.showwarning("Missing Info", "Please select your department and name.")
        return

    columns = ["Name", "Department", "Device Name", "Model", "Timestamp", "Action"]
    if os.path.exists(log_path):
        log_df = pd.read_excel(log_path)
    else:
        log_df = pd.DataFrame(columns=columns)

    if action == "Return":
        user_log = log_df[(log_df["Name"] == name) & (log_df["Device Name"] == device)]
        if user_log.empty or (user_log["Action"] == "Borrow").sum() <= (user_log["Action"] == "Return").sum():
            messagebox.showwarning("Invalid Return", f"No active borrow record for {device} by {name}.")
            return

    if action == "Borrow":
        device_log = log_df[log_df["Device Name"] == device]
        if (device_log["Action"] == "Borrow").sum() > (device_log["Action"] == "Return").sum():
            messagebox.showwarning("Device In Use", f"{device} is currently borrowed by someone else.")
            return

    new_row = pd.DataFrame([[name, dept, device, model, timestamp, action]], columns=columns)
    log_df = pd.concat([log_df, new_row], ignore_index=True)
    log_df.to_excel(log_path, index=False)
    messagebox.showinfo("Success", f"{action} recorded for {device} by {name}.")

# ---- UI Layout ----
# Borrow/Return Section
row_index = 0

# Department and employee dropdown for borrow/return
tk.Label(root, text="Department:", font=FONT).grid(row=row_index, column=0, padx=10, pady=5, sticky="e")
dept_dropdown = ttk.OptionMenu(root, dept_var, "", *dept_employee_map.keys())
dept_dropdown.grid(row=row_index, column=1, padx=10, pady=5, sticky="w")

row_index += 1
tk.Label(root, text="Employee:", font=FONT).grid(row=row_index, column=0, padx=10, pady=5, sticky="e")
emp_dropdown = ttk.OptionMenu(root, emp_var, "")
emp_dropdown.grid(row=row_index, column=1, padx=10, pady=5, sticky="w")

row_index += 1
tk.Label(root, text="Device Name:", font=FONT).grid(row=row_index, column=0, padx=10, pady=5, sticky="e")
device_dropdown = ttk.OptionMenu(root, device_name_var, *device_dict.keys())
device_dropdown.grid(row=row_index, column=1, padx=10, pady=5, sticky="w")

row_index += 1
tk.Label(root, text="Model:", font=FONT).grid(row=row_index, column=0, padx=10, pady=5, sticky="e")
tk.Entry(root, textvariable=model_var, font=FONT, state="readonly").grid(row=row_index, column=1, padx=10, pady=5, sticky="w")

row_index += 1
tk.Label(root, text="Vendor:", font=FONT).grid(row=row_index, column=0, padx=10, pady=5, sticky="e")
tk.Entry(root, textvariable=vendor_var, font=FONT, state="readonly").grid(row=row_index, column=1, padx=10, pady=5, sticky="w")

row_index += 1
tk.Label(root, text="Location:", font=FONT).grid(row=row_index, column=0, padx=10, pady=5, sticky="e")
tk.Entry(root, textvariable=location_var, font=FONT, state="readonly").grid(row=row_index, column=1, padx=10, pady=5, sticky="w")

row_index += 1
tk.Label(root, text="Status:", font=FONT).grid(row=row_index, column=0, padx=10, pady=5, sticky="e")
tk.Entry(root, textvariable=status_var, font=FONT, state="readonly").grid(row=row_index, column=1, padx=10, pady=5, sticky="w")

row_index += 1
tk.Label(root, text="Last Cleaned:", font=FONT).grid(row=row_index, column=0, padx=10, pady=5, sticky="e")
tk.Entry(root, textvariable=last_cleaned_var, font=FONT, state="readonly").grid(row=row_index, column=1, padx=10, pady=5, sticky="w")

row_index += 1
tk.Button(root, text="Borrow", command=lambda: log_action("Borrow"), bg="#4CAF50", fg="white", font=FONT).grid(row=row_index, column=0, padx=10, pady=10)
tk.Button(root, text="Return", command=lambda: log_action("Return"), bg="#2196F3", fg="white", font=FONT).grid(row=row_index, column=1, padx=10, pady=10)

# View borrowed equipment section further down
row_index += 2
tk.Label(root, text="--- View Borrowed Equipment ---", font=("Segoe UI", 12, "bold")).grid(row=row_index, column=0, columnspan=2, pady=10)

row_index += 1
tk.Label(root, text="Department:", font=FONT).grid(row=row_index, column=0, padx=10, pady=5, sticky="e")
view_dept_dropdown = ttk.OptionMenu(root, view_dept_var, "", *dept_employee_map.keys())
view_dept_dropdown.grid(row=row_index, column=1, padx=10, pady=5, sticky="w")

row_index += 1
tk.Label(root, text="Employee:", font=FONT).grid(row=row_index, column=0, padx=10, pady=5, sticky="e")
view_emp_dropdown = ttk.OptionMenu(root, view_emp_var, "")
view_emp_dropdown.grid(row=row_index, column=1, padx=10, pady=5, sticky="w")

row_index += 1
def show_borrowed_equipment():
    emp_name = view_emp_var.get().strip()
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

tk.Button(root, text="View Borrowed Equipment", command=show_borrowed_equipment,
          bg="#FF9800", fg="white", font=FONT).grid(row=row_index, column=0, columnspan=2, pady=10)

# Link dropdown behavior
dept_var.trace_add("write", lambda *args: update_employee_dropdown(dept_var, emp_var, emp_dropdown))
view_dept_var.trace_add("write", lambda *args: update_employee_dropdown(view_dept_var, view_emp_var, view_emp_dropdown))
device_name_var.trace_add("write", update_device_info)

# Initialize dropdowns
dept_var.set(list(dept_employee_map.keys())[0])
view_dept_var.set(list(dept_employee_map.keys())[0])
update_device_info()

root.mainloop()
