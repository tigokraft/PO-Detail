import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime, timedelta
import os

# Now the reference file must include these headers:
REF_REQUIRED_HEADERS = ["OPC", "On Hand", "ADD", "Descr", "SKU"]
# Input file still requires these:
INPUT_REQUIRED_HEADERS = ["OPC", "Due Date", "Ship Qty"]

def try_read_excel_line_check(filepath, required_headers):
    """
    Reads an Excel file. If the first row is empty or missing columns,
    tries again with the second row. Ensures 'required_headers' are found.
    """
    df = pd.read_excel(filepath, header=0)
    if df.iloc[0].isna().all():
        df = df.iloc[1:].reset_index(drop=True)

    if all(col in df.columns for col in required_headers):
        return df

    df2 = pd.read_excel(filepath, header=1)
    if df2.iloc[0].isna().all():
        df2 = df2.iloc[1:].reset_index(drop=True)

    if all(col in df2.columns for col in required_headers):
        return df2

    raise ValueError(
        f"Could not find required headers {required_headers} in '{filepath}' "
        "(checked header=0 and header=1)."
    )

def get_file_creation_date(filepath):
    """Gets the creation date of the file as a datetime."""
    return datetime.fromtimestamp(os.path.getctime(filepath))

# Format numeric columns
def format_add(value):
    try:
        s = str(value).replace(",", "").strip()
        if not s:
            return 0.0
        return float(s)
    except:
        return 0.0

def format_on_hand(value):
    try:
        s = str(value).replace(",", "").strip()
        if not s:
            return 0.0
        return float(s)
    except:
        return 0.0

def format_ship_qty(value):
    try:
        s = str(value).replace(",", "").strip()
        if not s:
            return 0.0
        return float(s)
    except:
        return 0.0

def process_inventory():
    if not input_file.get() or not reference_file.get() or not output_file.get():
        messagebox.showerror("Error", "Please select all files before running the process.")
        return

    input_filepath = input_file.get()
    reference_filepath = reference_file.get()
    output_filepath = output_file.get()

    try:
        df_input = try_read_excel_line_check(input_filepath, INPUT_REQUIRED_HEADERS)
    except Exception as e:
        messagebox.showerror("Error", f"Failed reading input file:\n{str(e)}")
        return

    try:
        df_ref = try_read_excel_line_check(reference_filepath, REF_REQUIRED_HEADERS)
    except Exception as e:
        messagebox.showerror("Error", f"Failed reading reference file:\n{str(e)}")
        return

    df_input["Ship Qty"] = df_input["Ship Qty"].apply(format_ship_qty)
    df_ref["On Hand"]    = df_ref["On Hand"].apply(format_on_hand)
    df_ref["ADD"]        = df_ref["ADD"].apply(format_add)

    df_input["Due Date"] = pd.to_datetime(
        df_input["Due Date"], 
        dayfirst=True, 
        errors="coerce"
    ).dt.strftime("%d/%m")
    df_input = df_input[df_input["Due Date"] != "NaT"]

    # Keep all OPCs from reference, even if they are not in input
    ref_opcs = set(df_ref["OPC"])
    df_input = df_input[df_input["OPC"].isin(ref_opcs)]  # Only keep OPCs that exist in reference
    df_ref["OPC"] = df_ref["OPC"].astype(str)

    # Ensure every OPC in reference appears in input by adding missing ones with Ship Qty = 0
    missing_opcs = ref_opcs - set(df_input["OPC"])
    if missing_opcs:
        missing_rows = pd.DataFrame({"OPC": list(missing_opcs), "Due Date": None, "Ship Qty": 0})
        df_input = pd.concat([df_input, missing_rows], ignore_index=True)

    file_creation_date = get_file_creation_date(input_filepath)
    date_columns = [(file_creation_date + timedelta(days=i)).strftime("%d/%m") for i in range(1, 9)]

    df_ship = df_input.groupby(["OPC", "Due Date"])["Ship Qty"].sum().reset_index()

    inventory_rows = []
    negative_rows = []

    for _, row in df_ref.iterrows():
        opc   = row["OPC"]
        descr = row["Descr"]
        sku   = row["SKU"]
        oh    = row["On Hand"]
        add   = row["ADD"]

        sub_df = df_ship[df_ship["OPC"] == opc]
        ship_dict = dict(zip(sub_df["Due Date"], sub_df["Ship Qty"]))

        current_stock = oh
        days_to_negative = None
        daily_inv = []

        for d_idx, day_str in enumerate(date_columns, start=1):
            if day_str in ship_dict:
                current_stock += ship_dict[day_str]
            current_stock -= add
            daily_inv.append(current_stock)

            if days_to_negative is None and current_stock < 0:
                days_to_negative = d_idx

        total_ship = sub_df["Ship Qty"].sum() if len(sub_df) else 0.0
        row_data = [descr, sku, opc, oh, add, total_ship] + daily_inv
        inventory_rows.append(row_data)

        if days_to_negative is not None and days_to_negative <= 4:
            negative_rows.append([descr, sku, opc, oh, add, days_to_negative])

    inv_cols = ["Description", "SKU", "OPC", "On Hand", "ADD", "Total Ship Qty"] + date_columns
    df_inv = pd.DataFrame(inventory_rows, columns=inv_cols)

    neg_cols = ["Description", "SKU", "OPC", "Initial On Hand", "Daily ADD", "Days to Negative"]
    df_neg = pd.DataFrame(negative_rows, columns=neg_cols)

    try:
        with pd.ExcelWriter(output_filepath, engine="openpyxl") as writer:
            df_inv.to_excel(writer, sheet_name="Inventory Status", index=False)
            df_neg.to_excel(writer, sheet_name="Negative Tracking", index=False)
        messagebox.showinfo("Success", f"Output saved to {output_filepath}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed writing output:\n{str(e)}")

# GUI
def browse_input():
    f = filedialog.askopenfilename(title="Select Input File", filetypes=[("Excel files","*.xlsx")])
    if f:
        input_file.set(f)

def browse_ref():
    f = filedialog.askopenfilename(title="Select Reference File", filetypes=[("Excel files","*.xlsx")])
    if f:
        reference_file.set(f)

def browse_out():
    f = filedialog.asksaveasfilename(title="Select Output File", defaultextension=".xlsx", filetypes=[("Excel files","*.xlsx")])
    if f:
        output_file.set(f)

root = tk.Tk()
root.title("Inventory Forecast: Descr & SKU in Negative Tracking")

input_file = tk.StringVar()
reference_file = tk.StringVar()
output_file = tk.StringVar()

tk.Label(root, text="Input File (Shipments):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
tk.Entry(root, textvariable=input_file, width=50).grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="Browse", command=browse_input).grid(row=0, column=2, padx=5, pady=5)

tk.Label(root, text="Reference File (Stock/ADD/Descr/SKU):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
tk.Entry(root, textvariable=reference_file, width=50).grid(row=1, column=1, padx=5, pady=5)
tk.Button(root, text="Browse", command=browse_ref).grid(row=1, column=2, padx=5, pady=5)

tk.Label(root, text="Output File:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
tk.Entry(root, textvariable=output_file, width=50).grid(row=2, column=1, padx=5, pady=5)
tk.Button(root, text="Browse", command=browse_out).grid(row=2, column=2, padx=5, pady=5)

tk.Button(root, text="Run Forecast", command=process_inventory, bg="green", fg="white").grid(row=3, column=1, pady=10)

root.mainloop()
