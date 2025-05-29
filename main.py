import pandas as pd
import os
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

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

root.mainloop()
