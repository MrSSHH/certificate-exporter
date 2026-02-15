import sqlite3
import json
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from tkinter.messagebox import showinfo
import pandas as pd


def select_file():
    filetypes = (('DB Files', '*.db'),)
    filename = filedialog.askopenfilename(title='Select DB', initialdir='.', filetypes=filetypes)
    if filename:
        showinfo(title='Selected File', message=filename)
        db_path.set(filename)
        update_dropdown()
        load_supplier_nms()


def load_tables():
    with open("tables.json", 'r') as json_file:
        db_tables = json.load(json_file)
        tables_names = list(db_tables.keys())
    return tables_names


def update_dropdown(*args):
    tables_names = load_tables()
    clicked.set(tables_names[0] if tables_names else '')
    drop['menu'].delete(0, 'end')
    for table in tables_names:
        drop['menu'].add_command(label=table, command=lambda value=table: clicked.set(value))


def load_supplier_nms():
    db_path_value = db_path.get()
    if db_path_value:
        conn = sqlite3.connect(db_path_value)
        cursor = conn.cursor()
        cursor.execute("SELECT NM, _id FROM SUPPLIER_ENTITY")
        supplier_data = cursor.fetchall()
        conn.close()

        global all_supplier_data
        all_supplier_data = supplier_data

        sorted_supplier_data = sorted(all_supplier_data, key=lambda x: x[0], reverse=False)
        update_supplier_dropdown(sorted_supplier_data)


def update_supplier_dropdown(supplier_data):
    for widget in root.grid_slaves():
        if isinstance(widget, OptionMenu):
            widget.destroy()

    supplier_dropdown = ttk.OptionMenu(root, supplier_var, '')
    supplier_dropdown.grid(row=4, column=1, padx=10, pady=10)

    for nm, supplier_id in supplier_data:
        supplier_dropdown['menu'].add_command(label=nm, command=lambda value=supplier_id: supplier_var.set(value))


def live_search(event):
    search_text = supplier_search_entry.get().lower()
    filtered_data = [item for item in all_supplier_data if search_text in item[0].lower()]
    update_supplier_dropdown(filtered_data)


def search_certificates_for_supplier():
    progress_bar.grid(row=9, column=0, columnspan=2, pady=10)
    progress_bar.start()

    supplier_id = supplier_var.get()
    db_path_value = db_path.get()

    if db_path_value and supplier_id:
        with open("tables.json", 'r') as json_file:
            db_tables = json.load(json_file)

        table_name = clicked.get()
        if table_name in db_tables:
            table_data = query_table_from_db(table_name)

            filtered_data = []
            for entry in table_data:
                if isinstance(entry, dict) and 'SUPPLIER_C' in entry and entry['SUPPLIER_C'] == int(supplier_id):
                    filtered_data.append(entry)

            if filtered_data:
                update_certificate_dropdown(filtered_data)
            else:
                showinfo("No Results", "No certificates found for the selected supplier _id.")
    

    progress_bar.stop()
    progress_bar.grid_forget()


def query_table_from_db(table_name):
    db_path_value = db_path.get()
    conn = sqlite3.connect(db_path_value)
    cursor = conn.cursor()

    cursor.execute(f"SELECT * FROM {table_name}")
    rows = cursor.fetchall()

    columns = [description[0] for description in cursor.description]
    table_data = []
    for row in rows:
        row_data = dict(zip(columns, row))
        table_data.append(row_data)

    conn.close()
    return table_data


def update_certificate_dropdown(certificates):
    # Clear existing options in the certificate dropdown
    menu = certificate_dropdown["menu"]
    menu.delete(0, 'end')

    # Store certificate data globally for later use in export
    global certificate_data_map
    certificate_data_map = {}

    for cert in certificates:
        cert_name = cert.get('SUPPLIER_NAME', 'Unknown Certificate')
        cert_id = cert.get('_id', 'Unknown ID')
        cert_date = cert.get('DATE', 'Unknown Date')
        
        # Store full certificate data keyed by ID (as string to match StringVar)
        certificate_data_map[str(cert_id)] = {
            'name': cert_name,
            'date': cert_date
        }
        
        menu.add_command(label=f"תעודה: {cert_name} - {cert_date}", command=lambda value=str(cert_id): certificate_var.set(value))

def export_related_data():
    certificate_id = certificate_var.get()
    with open('tables.json', 'r') as db_data:
        table_name = json.load(db_data)[clicked.get()]
    
    db_path_value = db_path.get()

    if db_path_value and certificate_id and table_name:
        table_data = query_table_from_db(table_name)
        matching_data = []

        for row in table_data:
            if int(row.get('PARENT_ENTITY_ID')) == int(certificate_id):
                # Append BARCODE and QUANTITY from matching rows
                matching_data.append({
                    "BARCODE": row.get('BARCODE'),
                    "QUANTITY": row.get('QUANTITY')
                })

        lines_to_take = int(lines_to_take_entry.get()) if lines_to_take_entry.get().isdigit() else len(matching_data)
        matching_data = matching_data[:lines_to_take]

        if matching_data:
            # Convert matching data to DataFrame with Hebrew column names
            df = pd.DataFrame(matching_data, columns=["BARCODE", "QUANTITY"])
            df.columns = ["ברקוד", "כמות"]  # Rename columns to Hebrew

            # Get supplier name and certificate date for filename
            supplier_name = "Unknown"
            cert_date = "Unknown"
            
            if certificate_id in certificate_data_map:
                supplier_name = certificate_data_map[certificate_id]['name']
                cert_date = certificate_data_map[certificate_id]['date']
            
            # Clean the supplier name and date for use in filename (remove invalid characters)
            supplier_name_clean = "".join(c for c in supplier_name if c.isalnum() or c in (' ', '_', '-')).strip()
            cert_date_clean = "".join(c for c in cert_date if c.isalnum() or c in (' ', '_', '-')).strip()
            
            # Define the file name with supplier name and date
            file_name = f"{supplier_name_clean}_{cert_date_clean}.xlsx"
            df.to_excel(file_name, index=False)

            # Show success message
            showinfo("Export Successful", f"Data exported to {file_name}")
        else:
            showinfo("No Results", "No matching rows found in the selected table.")


root = Tk()
root.title('Export Certificates')
root.geometry('800x500')
root.resizable(False, False)

style = ttk.Style()
style.theme_use("clam")
style.configure("TButton", padding=6, relief="flat", background="#4CAF50", foreground="white", font=('Arial', 10))
style.configure("TLabel", font=('Arial', 10), background="#f0f0f0")
style.configure("TOptionMenu", background="#f0f0f0", font=('Arial', 10))

db_path = StringVar()
supplier_var = StringVar()
certificate_var = StringVar()

# Global variable to store certificate data
certificate_data_map = {}

Label(root, text="Database file:").grid(row=0, column=0, padx=10, pady=10, sticky='e')
Label(root, textvariable=db_path).grid(row=0, column=1, padx=10, pady=10)

open_button = ttk.Button(root, text='Open a File', command=select_file)
open_button.grid(row=1, column=0, columnspan=2, pady=5)

Label(root, text="Select Table").grid(row=2, column=0, padx=10, pady=10, sticky='e')

clicked = StringVar()

tables_names = load_tables()
clicked.set(tables_names[0] if tables_names else "No Tables Available")

drop = ttk.OptionMenu(root, clicked, *tables_names)
drop.grid(row=2, column=1, padx=10, pady=10)

Label(root, text="Supplier Search").grid(row=3, column=0, padx=10, pady=10, sticky='e')
supplier_search_entry = Entry(root)
supplier_search_entry.grid(row=3, column=1, padx=10, pady=10)
supplier_search_entry.bind('<KeyRelease>', live_search)

Label(root, text="Supplier").grid(row=4, column=0, padx=10, pady=10, sticky='e')

supplier_dropdown = ttk.OptionMenu(root, supplier_var, '')
supplier_dropdown.grid(row=4, column=1, padx=10, pady=10)

search_button = ttk.Button(root, text="Search Certificates", command=search_certificates_for_supplier)
search_button.grid(row=5, column=0, columnspan=2, pady=10)

Label(root, text="Certificate").grid(row=6, column=0, padx=10, pady=10, sticky='e')

certificate_dropdown = ttk.OptionMenu(root, certificate_var, '')
certificate_dropdown.grid(row=6, column=1, padx=10, pady=10)

Label(root, text="Lines to Export (default 300)").grid(row=7, column=0, padx=10, pady=10, sticky='e')
lines_to_take_entry = Entry(root)
lines_to_take_entry.grid(row=7, column=1, padx=10, pady=10)

export_button = ttk.Button(root, text="Export Matching Items", command=export_related_data)
export_button.grid(row=8, column=0, columnspan=2, pady=10)

# Create a progress bar widget
progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="indeterminate")

db_path.trace("w", lambda *args: update_dropdown())

root.mainloop()