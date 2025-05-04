import os, time
import win32com.client as win32

def convert_specified_excel_sheet_to_pdf(folder_path,excel_files,sheet_name):
    """
    Converts a specified sheet from all Excel files (.xlsx, .xlsm) in a folder to PDFs without closing other open Excel files.
    Automatically updates external links in Excel files.
    
    :param sheet_name: Name of the sheet to convert.
    """
    # Try to attach to an existing Excel instance, otherwise create a new one
    try:
        excel = win32.GetActiveObject("Excel.Application")
        excel_was_already_running = True
        print("🔄 Attached to existing Excel instance.")
    except:
        excel = win32.Dispatch("Excel.Application")
        excel_was_already_running = False
        print("✅ Started new Excel instance.")

    excel.Visible = True  # Keep Excel visible

    try:
        for excel_file in excel_files:
            excel_path = os.path.join(folder_path, excel_file)
            print(f"\n📂 Processing File: {excel_file}")

            try:
                # Open the workbook with auto-update links
                wb = excel.Workbooks.Open(excel_path, UpdateLinks=3)  # 3 = Update all links

                # Suppress other pop-ups
                excel.DisplayAlerts = False

                # Check if the specified sheet exists
                sheet_names = [sheet.Name for sheet in wb.Sheets]

                if sheet_name not in sheet_names:
                    print(f"❌ Sheet '{sheet_name}' not found in '{excel_file}'. Skipping...")
                    wb.Close(False)
                    continue

                # Convert specified sheet to PDF (Workbook name only)
                pdf_name = f"{os.path.splitext(excel_file)[0]}.pdf"  # Removed sheet name from filename
                pdf_path = os.path.join(folder_path, pdf_name)
                
                print(f"🖨 Exporting '{sheet_name}' from '{excel_file}' to '{pdf_name}'...")

                # Ensure sheet exists and has content
                ws = wb.Sheets(sheet_name)
                if ws.UsedRange.Rows.Count == 1 and ws.UsedRange.Columns.Count == 1 and not ws.Cells(1, 1).Value:
                    print(f"⚠ Warning: '{sheet_name}' in '{excel_file}' is empty. Skipping...")
                    wb.Close(False)
                    continue

                # Export to PDF (Format = 0 for PDF)
                ws.ExportAsFixedFormat(0, pdf_path)

                # Confirm successful conversion
                if os.path.exists(pdf_path):
                    print(f"✅ Successfully saved: {pdf_name}")
                else:
                    print(f"❌ Failed to create PDF for '{excel_file}'.")

                # Close only the workbook opened by the script
                wb.Close(False)

            except Exception as e:
                print(f"❌ Error processing '{excel_file}': {e}")

    finally:
        # Restore default Excel settings
        excel.DisplayAlerts = True

        if not excel_was_already_running:
            excel.Quit()  # Quit only if we started Excel in this script
            print("🛑 Closed Excel instance.")

        print("\n🎉 Conversion process completed.")

first_run = True

def main(first_run):
    sheet_to_convert = "single page"  # Change this to the desired sheet name
    if first_run:
        folder_path = input("Enter Folder Location: ").strip()
        first_run = False
    else:
        folder_path = input("Enter Folder Location to Cont. or Press Enter to Exit: ").strip()
        if not folder_path.strip():
            print("Exiting...")
            time.sleep(.7)
            os.system("taskkill /f /im cmd.exe")
            return
            
    if os.path.exists(folder_path):
        excel_files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xlsm'))]
        if excel_files:
            convert_specified_excel_sheet_to_pdf(folder_path,excel_files,sheet_to_convert)
        else:
            print("⚠ No Excel files found in the folder.")
    else:
        print(f"❌ Error: The folder '{folder_path}' does not exist.")
    main(first_run)
    
def intro():
    print("Welcome to Lobo's Automation")
    print("Excel to PDF Converter by Nathan Lobo")

intro()          
main(first_run)