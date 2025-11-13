# 🧩 Handle Merged Cells in Excel using Python (`openpyxl` + `pandas`)

This project demonstrates how to **read Excel files containing merged cells** and convert them into a clean, tabular format using **Python**, **openpyxl**, and **pandas**.

---

## 📘 Overview

Excel files often contain merged cells that make direct data extraction difficult.  
This script:
1. Loads an Excel file using `openpyxl`
2. Detects merged cells and maps them to their original ranges
3. Extracts and fills merged cell values consistently across all merged rows/columns
4. Converts the processed data into a **pandas DataFrame** for analysis or export

---

## 🧠 Key Features

- ✅ Reads Excel files with merged cells
- ✅ Automatically fills merged cell values into individual rows
- ✅ Outputs a clean and consistent DataFrame
- ✅ Prevents duplicate rows during processing

---

## 🏗️ Project Structure

handle_merged_cells/
│
├── merged_cells_example.xlsx # Sample Excel file with merged cells
├── handle_merged_cells.ipynb # Jupyter Notebook / Python script
└── README.md # Project documentation (this file)