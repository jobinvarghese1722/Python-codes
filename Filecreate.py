import pandas as pd
from openpyxl import Workbook
from docx import Document
from fpdf import FPDF
import os

# Load the Excel file
excel_file = "C:\\filecreation-pythoncode\\ReadExcel.xlsx"  # Replace with your file path
output_dir = "C:/filecreation-pythoncode/outputfiles"  # Folder to save created files
os.makedirs(output_dir, exist_ok=True)  # Create the folder if it doesn't exist

# Function to create a text file
def create_text_file(filename, content):
    with open(filename, "w") as file:
        if pd.isna(content):  # Check if content is NaN or missing
            content = ""  # Replace with an empty string
        file.write(content)
    print(f"Created text file: {filename}")

# Function to create a PDF file
def create_pdf_file(filename, content):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    if pd.isna(content):  # Check if content is NaN or missing
        content = ""  # Replace with an empty string
    pdf.multi_cell(0, 10, content)
    pdf.output(filename)
    print(f"Created PDF file: {filename}")

# Function to create a Word file
def create_word_file(filename, content):
    doc = Document()
    # Convert content to string and handle NaN or missing values
    if pd.isna(content):  # Check if content is NaN or missing
        content = ""  # Replace with an empty string
    content = str(content)  # Convert to string
    for line in content.split("\n"):  # Handle multi-line content
        doc.add_paragraph(line)
    doc.save(filename)
    print(f"Created Word file: {filename}")

# Function to create an Excel file
def create_excel_file(filename, content):
    wb = Workbook()
    ws = wb.active

    ws.title = "Content"
    ws["A1"] = content
    if pd.isna(content):  # Check if content is NaN or missing
        content = ""  # Replace with an empty string
    wb.save(filename)
    print(f"Created Excel file: {filename}")

# Read the first sheet (Summary) to check Execution Status
first_sheet = pd.read_excel(excel_file, sheet_name=0)  # Assuming the first sheet is at index 0

# Debug: Print column names in the Summary sheet
print("Column names in the Summary sheet:", first_sheet.columns)

# Check if 'Execution Status' column exists
if 'Execution Status' not in first_sheet.columns:
    raise KeyError("Column 'Execution Status' not found in the Summary sheet. Available columns: " + str(first_sheet.columns))

# Iterate through the rows of the Summary sheet
for index, row in first_sheet.iterrows():
    execution_status = str(row["Execution Status"]).strip().lower()  # Normalize the status
    if execution_status != "yes":  # Skip if Execution Status is not "Yes"
        continue

    # Get the sheet name to process (from the "Test Scenario" column)
    sheet_name = row["Test Scenario"]
    if pd.isna(sheet_name):  # Skip if sheet name is missing
        print(f"Skipping row {index + 1}: Sheet Name is missing.")
        continue

    # Read the specified sheet
    try:
        sheet_data = pd.read_excel(excel_file, sheet_name=sheet_name)
        print(f"Processing sheet: {sheet_name}")
    except Exception as e:
        print(f"Error reading sheet '{sheet_name}': {e}")
        continue

    # Process the data from the specified sheet
    for sheet_index, sheet_row in sheet_data.iterrows():
        # Check if the row has "YES" in the Execution Status column
        if str(sheet_row["Execution Status"]).strip().upper() == "YES":
            # Append file extension dynamically based on Doc Type
            doc_type = str(sheet_row["Doc Type"]).lower()
            filename = os.path.join(output_dir, f"{sheet_row['Filename']}.{doc_type}")
            content = sheet_row["Content Inside"]

            # Create files based on Doc Type
            if doc_type == "txt":
                create_text_file(filename, content)
            elif doc_type == "pdf":
                create_pdf_file(filename, content)
            elif doc_type == "docx":
                create_word_file(filename, content)
            elif doc_type == "xlsx":
                create_excel_file(filename, content)
            else:
                print(f"Unsupported document type: {doc_type} for file {filename}")

print(f"Files created successfully in the '{output_dir}' directory.")