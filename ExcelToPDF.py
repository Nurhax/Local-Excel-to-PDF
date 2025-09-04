import os
import win32com.client

# Function to clear filters from a sheet
def clear_filters(sheet):
    # If AutoFilter is on, clear it
    if sheet.AutoFilterMode:
        sheet.AutoFilterMode = False
    if sheet.FilterMode:
        try:
            sheet.ShowAllData()
        except:
            pass  # ShowAllData might fail if no filter is applied

# Function to hide columns after AU
def hide_columns_after_au(sheet):
    max_col = sheet.UsedRange.Columns.Count
    limit_col = 47  # AU is column 47

    for col in range(limit_col + 1, max_col + 1):
        sheet.Columns(col).Hidden = True

# Main function to convert Excel files to PDF
def convert_excel_to_pdf(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    for file_name in os.listdir(input_folder):
        if file_name.endswith((".xlsx", ".xls")):
            full_path = os.path.join(input_folder, file_name)
            workbook = excel.Workbooks.Open(full_path)

            try:
                for sheet in workbook.Sheets:
                    clear_filters(sheet)
                    hide_columns_after_au(sheet)

                output_path = os.path.join(
                    output_folder,
                    os.path.splitext(file_name)[0] + ".pdf"
                )

                workbook.ExportAsFixedFormat(0, output_path)
                print(f"✔️ Converted with borders: {file_name}")

            except Exception as e:
                print(f"❌ Failed to process {file_name}: {e}")
            finally:
                workbook.Close(False)

    excel.Quit()
    print("✅ All files processed and exported as PDFs with filters removed and borders applied.")

# Input Output Path Folders
input_folder = r"C:\KP\Files\PLNSHIZ\ExcelIFS"
output_folder = r"C:\KP\Files\PLNSHIZ\PDFs"

convert_excel_to_pdf(input_folder, output_folder)
