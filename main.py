import os
import glob
import win32print
import win32com.client
import re  # Import the re module for regular expressions
import tkinter as tk
from tkinter import messagebox

# Define a password for the script
PASSWORD = "sunita"  # Replace with your desired password

def ask_for_password():
    """
    Ask the user for the password.
    Returns True if the password is correct, False otherwise.
    """
    entered_password = input("Enter password to continue: ")
    return entered_password == PASSWORD

def is_printer_online(printer_name):
    """
    Check if the printer is online and ready.
    Returns True if the printer is ready, False if it is offline or has errors.
    """
    try:
        # Open the printer to get a valid printer handle
        printer_handle = win32print.OpenPrinter(printer_name)
        
        # Get printer info using the valid handle (we pass None as pServerName for local printers)
        printer_info = win32print.GetPrinter(printer_handle, 2)
    
        # Close the printer handle after use
        win32print.ClosePrinter(printer_handle)
        
        # Check the attributes and ensure the printer is both online and ready
        attributes = printer_info['Attributes']
        
        # Printer is ready if attributes match the expected value (indicating it's online)
        if attributes == 2624:  # Printer is ready
            return True
        else:
            return False

    except Exception as e:
        print(f"Error checking printer status: {e}")
        return False

def extract_date_from_filename(filename):
    """
    Extract the date (MM-DD-YYYY) from filenames like 'RDInstallmentReport03-09-2024'.
    Returns the date as a string or None if the date is not found.
    """
    pattern = r"RDInstallmentReport(\d{2}-\d{2}-\d{4})"
    match = re.search(pattern, filename)
    if match:
        return match.group(1)  # Return the date part (MM-DD-YYYY)
    return None

def filter_excel_files(excel_files):
    """
    Filters Excel files to match those with the pattern 'RDInstallmentReport' followed by a date.
    Returns a dictionary with dates as keys and lists of files as values.
    """
    date_files_map = {}
    for file in excel_files:
        date = extract_date_from_filename(os.path.basename(file))
        if date:
            if date not in date_files_map:
                date_files_map[date] = []
            date_files_map[date].append(file)
    return date_files_map

def print_excel_files_from_folder(excel_files):
    if not excel_files:
        print("No Excel files found in the current directory.")
        return
    
    # Initialize Excel COM object
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False  # Keep Excel invisible during operation

    # Get the default printer
    printer_name = win32print.GetDefaultPrinter()

    print(f"Default Printer: {printer_name}")  # Print the printer name for debugging
    
    # Check if the printer is online before proceeding
    if not is_printer_online(printer_name):
        print(f"Printer '{printer_name}' is not ready. OFFLINE")
        excel.Quit()
        return

    for file in excel_files:
        try:
            # Open the Excel file
            print(f"Printing {file}...")
            wb = excel.Workbooks.Open(file)
            
            # Get the first sheet of the workbook
            sheet = wb.Sheets(1)  # You can change the index if you want a different sheet
            
            # Format a specific cell (e.g., cell I6)
            cell = sheet.Range("I6")  # Change this to the specific cell you want to format
            cell.Font.Bold = True      # Make the font bold
            cell.Font.Size = 19        # Set the font size

            # Remove any print area that might be set
            sheet.PageSetup.PrintArea = ""

            # Set the paper size to A4 (using numeric value)
            sheet.PageSetup.PaperSize = 9  # A4 paper size

            # Set the page orientation to Portrait (if necessary)
            sheet.PageSetup.Orientation = 1  # 1 is Portrait, 2 is Landscape

            # Set the sheet to fit to one page (both horizontally and vertically)
            sheet.PageSetup.FitToPagesWide = 1  # Fit to 1 page wide
            sheet.PageSetup.FitToPagesTall = 1  # Fit to 1 page tall

            # Optional: Adjust margins to make the page fit better
            sheet.PageSetup.LeftMargin = excel.Application.InchesToPoints(0.5)
            sheet.PageSetup.RightMargin = excel.Application.InchesToPoints(0.5)
            sheet.PageSetup.TopMargin = excel.Application.InchesToPoints(0.5)
            sheet.PageSetup.BottomMargin = excel.Application.InchesToPoints(0.5)

            # Optionally force zoom to ensure the content fits
            sheet.PageSetup.Zoom = False  # Disable the Zoom property
            sheet.PageSetup.FitToPagesWide = 1  # Ensure fitting horizontally
            sheet.PageSetup.FitToPagesTall = 1  # Ensure fitting vertically

            # Print the active sheet
            sheet.PrintOut()

            # Close the workbook after printing
            wb.Close(False)
            
        except Exception as e:
            print(f"Failed to print {file}: {e}")

    # Quit Excel application after processing all files
    excel.Quit()

def ask_user_to_select_dates(date_files_map):
    """
    Displays a Tkinter window with checkboxes for the user to select the dates they want to print.
    Returns a list of selected dates.
    """
    root = tk.Tk()
    root.title("Select Dates to Print")

    # Create a list to store the selected dates
    selected_dates = []

    # Function to toggle selection of a date
    def toggle_date_selection(date, var):
        if var.get():
            selected_dates.append(date)
        else:
            selected_dates.remove(date)

    # Create a checkbox for each date
    for date in date_files_map.keys():
        var = tk.BooleanVar()
        checkbox = tk.Checkbutton(root, text=date, variable=var)
        checkbox.pack(anchor='w')
        checkbox.config(command=lambda date=date, var=var: toggle_date_selection(date, var))

    # Done button to close the window
    done_button = tk.Button(root, text="Done", command=root.quit)
    done_button.pack()

    # Run the Tkinter event loop
    root.mainloop()

    return selected_dates

# Main program logic

# Ask the user for password before running the script
if not ask_for_password():
    print("Incorrect password. Exiting the script.")
else:
    # Get the current directory where the script is executed
    folder_path = os.getcwd()  # Get the current working directory

    # Use glob to find all Excel files (.xlsx, .xls) in the current directory
    excel_files = glob.glob(os.path.join(folder_path, '*.xlsx')) + glob.glob(os.path.join(folder_path, '*.xls'))

    # Filter the Excel files based on the pattern and extract dates
    date_files_map = filter_excel_files(excel_files)

    if date_files_map:
        # Ask the user to select the dates to print
        selected_dates = ask_user_to_select_dates(date_files_map)

        if selected_dates:
            # Get the files corresponding to the selected dates
            files_to_print = [file for date in selected_dates for file in date_files_map[date]]
            print_excel_files_from_folder(files_to_print)
        else:
            print("No dates selected to print.")
    else:
        print("No valid files found in the directory.")
