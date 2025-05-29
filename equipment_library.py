import pandas as pd
import os
from tkinter import messagebox


def load_inventory(inventory_path):
    if not os.path.exists(inventory_path):
        messagebox.showerror("Missing File", f"{inventory_path} not found.")
        raise FileNotFoundError(f"{inventory_path} not found.")

    inventory_df = pd.read_excel(inventory_path)
    inventory_df = inventory_df.fillna("N/A")

    device_dict = {}
    for _, row in inventory_df.iterrows():
        key = f"{row['DESCRIPTION']} ({row['EQUIPMENT TYPE']})"
        device_dict[key] = {
            "model": str(row["ID CODE/ MODEL NO."] or "N/A"),
            "vendor": row["VENDOR"],
            "location": row["STOCK LOCATION"],
            "status": row["STATUS"],
            "last_cleaned": str(row["LAST CLEAN DATE"]),
            "number_in_stock": row["NUMBER IN STOCK"]
        }

    return device_dict


def get_department_employee_map():
    return {
        "Manager/Supervisor": [
            "Phuong Hoang", "Phat Dinh", "Hao Cao", "Trang Huynh", "Ngan Duong", "Dung Huynh",
            "Buu Phan", "Edward Bui", "Edana Nguyen", "Ha Doan", "Thuy Le"
        ],
        "Research & Development": ["Linh Tran"],
        "Production": [
            "Vi Dam", "Tuyen Do", "Tuan Tran", "Hoang Thai", "Vinh Dao", "Trong Nguyen", "Hoang Le",
            "Thinh Pham", "Kim Anh Dang", "Thuy Tran", "Sinh Pham", "Nam Huynh", "Hien Nguyen"
        ],
        "Sales": ["Tina Truong", "Thanh Dang", "Duc Nguyen"],
        "Retails": ["Hiep Tran", "Thuy Nguyen"],
        "Customer Service": ["Huong Pham", "Nhi Dam", "Vu Tran"],
        "Marketing": [
            "Hang Nguyen", "Phuong Pham", "Thu Vo", "Nien Tran", "Huyen Ho", "Tuyen Ho", "Bao Chau", "Loan Pham"
        ],
        "Inventory": ["Alec Hall", "Huy Kieu", "Vuong Nguyen"],
        "Accounting": ["Huy Nguyen", "Trinh Nguyen", "Tami Nguyen"],
        "IT": ["An Vuong", "Khang Truong", "Dung Nguyen", "Quang Pham", "Vinh Pham"],
        "Shipping": [
            "Man Le", "Tan Nguyen", "Thuy Bui", "Xuan Luong", "Binh Dang", "Tien Chau", "Dat Chau",
            "Hien Duong", "Tai Nguyen"
        ]
    }
