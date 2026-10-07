# Excel Field Extractor & Sorter

A Python desktop GUI application built with Tkinter for easily extracting, filtering, and sorting specific columns/fields from Excel files.

## Features
- **Upload Excel files**: Supports both `.xlsx` and `.xls` formats.
- **Preview Data**: Instantly view the original Excel data in a grid after uploading.
- **Select Sheet**: If an Excel workbook has multiple sheets, you can select which one to process.
- **Search & Filter Columns**: Dynamically search for columns/fields to easily find what you need.
- **Multiple Selection**: Select one or multiple fields to fetch.
- **Alphabetical Sorting**: Automatically sorts fetched data alphabetically (A-Z or Z-A) and removes empty/null values.
- **Export Data**: Save the newly filtered and sorted result into a new `.xlsx` file.

## Requirements
- Python 3.7+
- pandas
- openpyxl
- xlrd

## Installation & Setup

1. **Clone or Extract the Project**
   Extract the files to your desired folder.

2. **Create a Virtual Environment (Recommended)**
   Open a terminal/command prompt in the project folder and run:
   ```bash
   python -m venv venv
   ```

3. **Activate the Virtual Environment**
   - Windows:
     ```bash
     venv\Scripts\activate
     ```
   - macOS/Linux:
     ```bash
     source venv/bin/activate
     ```

4. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

## How to Run

Execute the `main.py` script from your command prompt or terminal:
```bash
python main.py
```

## Usage Workflow
1. Click **Upload Excel File** and choose your document.
2. Select the target **Sheet** from the dropdown menu (if multiple exist).
3. (Optional) Use the **Search Field** box to find specific columns.
4. Click on the fields in the list to select them (they will be highlighted).
5. Choose your preferred sorting order (**A → Z** or **Z → A**).
6. Click **Fetch Selected Fields** to view the sorted results in the "Filtered Result" tab.
7. Click **Export to Excel** to save the result.
