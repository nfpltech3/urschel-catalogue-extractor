# Urschel PDF Generator - User Guide

## Introduction
The Urschel PDF Generator is an internal, automated desktop utility designed to rapidly scan massive machinery OEM PDF catalogues and automatically extract the specific pages containing product models listed in an Excel sheet. The final output is a clean, consolidated PDF document containing only the requested information.

## How to Use

### 1. Launching the App
Simply double-click the `Urschel_Catalogue_Tool.exe` file. The application will launch in fullscreen mode automatically.
*(The tool requires no database connections or API keys to operate)*

### 2. The Workflow (Step-by-Step)
1. **Browse Excel...**: Click the button under **Step 1** to select your Excel file containing the list of required parts.
   - *Note: The tool strictly accepts `.xlsx` or `.xls` files. Inside the file, there MUST be a column named exactly `Model` containing the part numbers.*
2. **Select PDF Catalogues...**: Click the button under **Step 2** to select the large source PDF catalogue(s). You can select multiple PDFs at once.
   - *Note: The tool remembers the catalogues you chose in your last session and will auto-load them for convenience the next time you open the app.*
3. **Number of models to fetch**: Enter the numerical limit of how many models you want to extract into the input field in **Step 3**.
   - *Note: To ensure rapid speed, the scanner will stop completely once this target is hit. If you want to scan all models, simply enter a number higher than the total rows in your Excel file.*
4. **GENERATE PDF**: Click this primary button to begin extraction.
   - *Note: The tool will flash "Preparing PDFs..." as it caches the text of the giant files into memory. Once cached, it extracts the pages near-instantly.*

## Interface Reference

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| **Browse Excel...** | Opens a file dialog to select the mapping sheet | File format: `.xlsx`, `.xls` |
| **Select PDF Catalogues...** | Opens a file dialog to select the massive source manuals | File format: `.pdf` |
| **Number of models to fetch** | Number input to cap extraction limits and maximize speed. | A whole Integer (e.g., `10`) |
| **GENERATE PDF** | Triggers the actual scanning and extracting mechanism. | N/A |
| **Data Logs Table** | Displays the real-time extraction results including the Model, Status, PDF Source, and Page Numbers. | N/A |

## Troubleshooting & Validations

If you see an error or experience unexpected behavior, check this table:

| Message / Behavior | What it means | Solution |
| :--- | :--- | :--- |
| **"Please select the Excel file and at least one PDF."** | You clicked GENERATE without completing Steps 1 and 2. | Select your data files using the Browse buttons. |
| **"An error occurred: KeyError: 'Model'"** | The Python script could not find the required column in the Excel file. | Open your Excel file and rename the column header containing your part numbers to exactly `Model`. |
| **Your Excel file does not appear in the "Browse" list** | Windows is filtering for actual Excel files, but your file is likely a `.csv` despite the green icon. | Open the file, click "Save As", and change the type to "Excel Workbook (\*.xlsx)". |
| **A model says "Not Found" in the Logs Table** | The scanner explicitly checked every page of every selected PDF and could not find that part number. | Check if the Catalogue PDF containing that specific part was actually selected in Step 2. |
| **"No matching models were found."** | Out of the entire Excel list, zero items were located in the PDFs. | Verify the part numbers are formatted correctly with no bizarre spacing differences compared to the PDF text. |
