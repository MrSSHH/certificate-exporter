# 📊 SQLite Certificate Exporter

A professional Python-based GUI utility designed to bridge the gap between complex SQLite databases and user-friendly Excel reporting. This tool allows users to navigate supplier entities, filter associated certificates, and export relational data (like barcodes and quantities) with localized Hebrew headers.

---

## 📝 Project Overview

This application serves as a **Search-and-Export bridge** for inventory and logistics management.  

Instead of writing manual SQL queries to locate items within a specific delivery note or certificate, users can utilize a visual interface to filter, preview, and generate reports in seconds.

---

## ✨ Key Features

- 🔎 **Live Supplier Search** – Real-time filtering of supplier lists as you type.
- 🗂️ **Dynamic Table Mapping** – Uses a `tables.json` file to map user-friendly names to actual SQL table names.
- 📄 **Localized Excel Export** – Automatically generates `.xlsx` files with Hebrew column headers (`ברקוד` and `כמות`).
- 🏷️ **Smart Filenaming** – Exports are named based on the selected supplier and certificate date for easy organization.
- ⏳ **Progress Feedback** – Includes an indeterminate progress bar for database-heavy operations.

---

## 🛠️ Installation & Setup

### 1️⃣ Prerequisites

- Python **3.8 or higher**

Check your Python version:

```bash
python --version
````

---

### 2️⃣ Install Dependencies

This project relies on external libraries for data processing and Excel generation.

Install required packages:

```bash
pip install -r requirements.txt
```

---

### 3️⃣ Repository Files

| File               | Description                                               |
| ------------------ | --------------------------------------------------------- |
| `main.py`          | Core application code                                     |
| `requirements.txt` | List of required Python libraries                         |
| `tables.json`      | Configuration file for database schema                    |
| `.gitignore`       | Prevents temporary files and databases from being tracked |

---

## ⚙️ Configuration (`tables.json`)

The application requires a `tables.json` file in the root directory to determine which tables to query.

Example format:

```json
{
    "ENTRY_APPROVAL_CERTIFICATE_ENTITY": "ENTRY_APPROVAL_ITEM_ENTITY",
    "ENTRY_CERTIFICATE_ENTITY": "ENTRY_CERTIFICATE_ITEM_ENTITY",
    "REFUND_CERTIFICATE_ENTITY": "REFUND_CERTIFICATE_ITEM_ENTITY",
    "STOCK_TAKING_ENTITY": "STOCK_TAKING_ITEM_ENTITY",
}
```

---

## 📖 How to Use

1. **Launch the App**

   ```bash
   python main.py
   ```

2. **Select Database**

   * Click **Open a File**
   * Choose your SQLite `.db` file

3. **Choose Table**

   * Select a table from the dropdown menu (populated via `tables.json`)

4. **Find Supplier**

   * Start typing in the **Supplier Search** box
   * The dropdown updates in real time

5. **Fetch Certificates**

   * Click **Search Certificates** to retrieve related entries

6. **Export Data**

   * Select a certificate
   * Set row limit (default: `300`)
   * Click **Export Matching Items**

---



