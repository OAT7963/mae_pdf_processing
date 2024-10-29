import tkinter as tk
from tkinter import filedialog, ttk  # ttk for improved widgets
from tkinter import messagebox
import fitz  # PyMuPDF
import pandas as pd
import re
import os

# List of strings to remove during processing
strings_to_remove = [
    "URUSNIAGA AKAUN/ 戶口進支項 /ACCOUNT TRANSACTIONS",
    "TARIKH MASUK",
    "BUTIR URUSNIAGA",
    "JUMLAH URUSNIAGA",
    "BAKI PENYATA",
    "進支日期",
    "進支項說明",
    "银碼",
    "結單存餘",
    "URUSNIAGA AKAUN/ 戶口進支項/ACCOUNT TRANSACTIONS",
    "TARIKH NILAI",
    "仄過賬日期"
]

# Improved directory selection row creation
def create_directory_selection_row(root, label_text, browse_command, entry_width=30, row=0):
    label = ttk.Label(root, text=label_text, background='white')
    label.grid(row=row, column=0, sticky=tk.W, padx=(10, 5), pady=(5, 5))

    entry = ttk.Entry(root, width=entry_width)
    entry.grid(row=row, column=1, sticky=tk.EW, padx=(0, 5), pady=(5, 5))

    button = ttk.Button(root, text="Browse", command=lambda: browse_command(entry))
    button.grid(row=row, column=2, padx=(5, 10), pady=(5, 5))

    root.grid_columnconfigure(1, weight=1)  # This makes the entry expand to fill the column

    return entry

def selected_processing():
    mode = processing_mode.get()
    if mode == "Maybank Debit Card Statement Processing":
        process_files()
    elif mode == "Maybank Credit Card Statement Processing":
        process_file_cc_statement()
    elif mode == "CIMB Debit Statement Processing":
        process_CIMB_DEBIT_data()
    elif mode == "M2U Current Account Statement":
        process_files_m2u()
    elif mode == "M2U Current Account Debit":
        process_files_m2u_debit()
    else:
        messagebox.showerror("Error", "Invalid processing mode selected")

def select_directory(entry):
    folder_path = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, folder_path)

def remove_sections(lines, start_marker, end_marker):
    new_lines = []
    in_section = False
    for line in lines:
        if start_marker in line:
            in_section = True
            continue  # Skip adding this line
        elif end_marker in line:
            in_section = False
            continue  # Skip adding this line and move past the end marker
        if not in_section:
            new_lines.append(line)
    return new_lines

def determine_flow(transaction_amount):
    if transaction_amount.endswith('+'):
        return 'Deposit'
    elif transaction_amount.endswith('-'):
        return 'Withdrawal'
    else:
        return 'unknown'

def process_m2u_statement(pdf_path):
    # Read PDF
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()

    # Split into lines
    lines = text.split('\n')

    # Extract year from statement date
    date_pattern = re.compile(r'\d{2}/\d{2}/\d{2}')
    year_statement = "00"
    for line in lines:
        if date_pattern.match(line):
            year_match = re.match(r'(\d{2})/(\d{2})/(\d{2})', line)
            if year_match:
                year_statement = year_match.group(3)
            break

    # Remove unnecessary sections
    lines = remove_sections(lines, 'Malayan Banking Berhad (3813-K)', 'denoted by DR')
    lines = remove_sections(lines, 'FCN', 'PLEASE BE INFORMED TO CHECK YOUR BANK ACCOUNT BALANCES REGULARLY')
    lines = remove_sections(lines, 'ENTRY DATE', 'STATEMENT BALANCE')
    lines = remove_sections(lines, 'ENDING BALANCE :', 'TOTAL CREDIT :')

    # Filter lines
    filtered_lines = [line for line in lines if not any(s in line for s in strings_to_remove)]
    
    # Process transactions
    date_pattern = re.compile(r'\d{2}/\d{2}')
    structured_data = []
    temp_entry = {}

    for line in filtered_lines:
        if date_pattern.match(line):
            if temp_entry:
                structured_data.append(temp_entry)
            temp_entry = {
                "Entry Date": line,
                "Transaction Description": "",
                "Transaction Amount": "",
                "Statement Balance": ""
            }
        elif "Transaction Amount" in temp_entry and temp_entry["Transaction Amount"] and "Statement Balance" in temp_entry and temp_entry["Statement Balance"] == "":
            temp_entry["Statement Balance"] = line.strip()
        elif "Transaction Amount" in temp_entry and temp_entry["Transaction Amount"] == "":
            temp_entry["Transaction Amount"] = line.strip()
        else:
            if temp_entry:
                temp_entry["Transaction Description"] += line.strip() + ", "

    if temp_entry:
        structured_data.append(temp_entry)

    # Convert to DataFrame
    df = pd.DataFrame(structured_data)

    # Process dates and add year
    df['Entry Date'] = pd.to_datetime(df['Entry Date'], format='%d/%m', dayfirst=True).dt.date
    df['Entry Date'] = df['Entry Date'].apply(lambda x: x.replace(year=2000 + int(year_statement)))

    # Process balances and amounts
    df['Statement Balance 2'] = df['Transaction Description'].str.extract(r'(\d+,\d+\.\d+)')[0]
    df['Statement Balance 2'] = df['Statement Balance 2'].str.replace(',', '').astype(float)

    # Clean description
    df['Transaction Description'] = df['Transaction Description'].str.replace(r'\d+,\d+\.\d+, ', '', regex=True)
    df['Transaction Description'] = df['Transaction Description'].str.replace(r', (\d{1,3}(?:,\d{3})*(?:\.\d{2}))$', '', regex=True)

    # Reorder and rename columns
    df = df[['Entry Date', 'Transaction Amount', 'Transaction Description', 'Statement Balance', 'Statement Balance 2']]
    df = df.rename(columns={
        'Transaction Amount': 'Transaction Type',
        'Statement Balance': 'Transaction Amount',
        'Statement Balance 2': 'Statement_Balance'
    })

    # Update specific transaction descriptions
    transaction_type_mappings = {
        'CASH WITHDRAWAL': 'CASH WITHDRAWAL',
        'DEBIT ADVICE': 'Card Annual Fee',
        'PROFIT PAID': 'PROFIT PAID',
        'INTEREST PAYMENT': 'INTEREST PAYMENT',
        'INT ON INT PAYMENT': 'INT ON INT PAYMENT'
    }
    
    for trans_type, description in transaction_type_mappings.items():
        df.loc[df['Transaction Type'] == trans_type, 'Transaction Description'] = description

    # Add flow and clean amount
    df['flow'] = df['Transaction Amount'].apply(determine_flow)
    df['Transaction Amount'] = df['Transaction Amount'].str.replace('+', '', regex=False).str.replace('-', '', regex=False)
    df['Transaction Amount'] = df['Transaction Amount'].str.replace(',', '').astype(float)

    return df

def process_files_m2u_debit():
    folder_path = source_path_entry.get()
    export_path = export_path_entry.get()
    excel_file_name = excel_name_entry.get()
    
    if not folder_path or not excel_file_name or not export_path:
        messagebox.showerror("Error", "Folder path, export path, or Excel file name is missing")
        return
    
    pdf_files = [file for file in os.listdir(folder_path) if file.endswith('.pdf')]
    all_data = []

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        try:
            # Process the statement using process_m2u_statement()
            df = process_m2u_statement(pdf_path)
            all_data.append(df)
        except Exception as e:
            print(f"Error processing {pdf_file}: {str(e)}")
            continue

    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        excel_path = os.path.join(export_path, f"{excel_file_name}.csv")
        combined_df.to_csv(excel_path, index=False)
        print(f"Data exported to {excel_path}")
        messagebox.showinfo("Success", f"Data exported successfully to {excel_path}")
    else:
        print("No data to export.")
        messagebox.showinfo("No Data", "No data was processed.")

def process_file_cc_statement():
    folder_path = source_path_entry.get()
    export_path = export_path_entry.get()
    excel_file_name = excel_name_entry.get()

    if not folder_path or not export_path or not excel_file_name:
        messagebox.showerror("Error", "Folder path, export path, or Excel file name is missing")
        return

    pdf_files = [file for file in os.listdir(folder_path) if file.endswith('.pdf')]
    all_data = []

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        try:
            doc = fitz.open(pdf_path)
            text = ""
            for page in doc:
                text += page.get_text()
            doc.close()
        except Exception as e:
            print(f"Error opening {pdf_path}: {e}")
            continue  # Skip to the next file

        # Extract the year from the filename
        pattern = re.compile(r"\d{4}")
        yearlist = pattern.findall(pdf_file)
        years = "0000"

        for years in yearlist:
            if ((int(years) > 2010) and (int(years) < 2050)):
                year = years

        lines = text.split('\n')
        filtered_lines = [line for line in lines if not any(s in line for s in strings_to_remove)]
        data = filtered_lines

        final_structured_data = []
        i = 0
        while i < len(data):
            if '/' in data[i] and len(data[i]) == 5 and '/' in data[i + 1] and len(data[i + 1]) == 5:
                transaction_date = data[i]
                posting_date = data[i + 1]
                i += 2
                
                description = []
                amount = ''
                
                while i < len(data) and not ('/' in data[i] and len(data[i]) == 5):
                    clean_line = data[i].strip()
                    amount_match = re.match(r'^(\d{1,3}(?:,\d{3})*(\.\d{2})?)(CR)?$', clean_line, re.IGNORECASE)
                    if amount_match:
                        amount = amount_match.group(1)
                        if amount_match.group(3):
                            amount = '-' + amount
                        i += 1
                        break
                    else:
                        description.append(clean_line)
                    i += 1
                
                description_text = ', '.join(description)
                
                final_structured_data.append([transaction_date, posting_date, description_text, amount, year])
            else:
                i += 1

        if final_structured_data:
            final_df = pd.DataFrame(final_structured_data, columns=['Posting Date', 'Transaction Date', 'Transaction Description', 'Amount', 'Year'])
            all_data.append(final_df)

    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        combined_df['Amount'] = combined_df['Amount'].str.replace(',', '').replace('', None).astype(float)
        combined_df['Year'] = combined_df['Year'].astype('Int64')
        combined_df = combined_df[['Year', 'Posting Date', 'Transaction Date', 'Transaction Description', 'Amount']]

        excel_path = os.path.join(export_path, f"{excel_file_name}.csv")
        combined_df.to_csv(excel_path, index=False)
        print(f"Data exported to {excel_path}")
        messagebox.showinfo("Success", f"Data exported successfully to {excel_path}")
    else:
        print("No data to export.")

def process_files_m2u():
    folder_path = source_path_entry.get()
    export_path = export_path_entry.get()
    excel_file_name = excel_name_entry.get()
    if not folder_path or not excel_file_name or not export_path:
        messagebox.showerror("Error", "Folder path, export path, or Excel file name is missing")
        return
    
    pdf_files = [file for file in os.listdir(folder_path) if file.endswith('.pdf')]
    all_data = []

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()

        lines = text.split('\n')

        date_pattern = re.compile(r'\d{2}/\d{2}/\d{2}')
        year_statement = "00"
        for line in lines:
            if date_pattern.match(line):
                year_match = re.match(r'(\d{2})/(\d{2})/(\d{2})', line)
                if year_match:
                    year_statement = year_match.group(3)
                break
        
        lines = remove_sections(lines, 'Malayan Banking Berhad (3813-K)', 'denoted by DR')
        lines = remove_sections(lines, 'FCN', 'PLEASE BE INFORMED TO CHECK YOUR BANK ACCOUNT BALANCES REGULARLY')
        lines = remove_sections(lines, 'ENTRY DATE', 'STATEMENT BALANCE')
        lines = remove_sections(lines, 'ENDING BALANCE :', 'TOTAL CREDIT :')

        filtered_lines = [line for line in lines if not any(s in line for s in strings_to_remove)]
        transactions = filtered_lines
        structured_data = []
        temp_entry = {}

        date_pattern = re.compile(r'\d{2}/\d{2}')

        for line in transactions:
            if date_pattern.match(line):
                if temp_entry:
                    structured_data.append(temp_entry)
                temp_entry = {"Entry Date": line, "Transaction Description": "", "Transaction Amount": "", "Statement Balance": ""}
            elif "Transaction Amount" in temp_entry and temp_entry["Transaction Amount"] and "Statement Balance" in temp_entry and temp_entry["Statement Balance"] == "":
                temp_entry["Statement Balance"] = line.strip()
            elif "Transaction Amount" in temp_entry and temp_entry["Transaction Amount"] == "":
                temp_entry["Transaction Amount"] = line.strip()
            else:
                if temp_entry:
                    temp_entry["Transaction Description"] += line.strip() + ", "

        if temp_entry:
            structured_data.append(temp_entry)

        for entry in structured_data:
            entry["Transaction Description"] = entry["Transaction Description"].rstrip(', ')

        df = pd.DataFrame(structured_data)
        df['Entry Date'] = pd.to_datetime(df['Entry Date'], format='%d/%m', dayfirst=True).dt.date 
        df['Entry Date'] = df['Entry Date'].apply(lambda x: x.replace(year = 2000 + int(year_statement)))
 
        df['Statement Balance 2'] = df['Transaction Description'].str.extract(r'(\d+,\d+\.\d+)')[0]
        df['Statement Balance 2'] = df['Statement Balance 2'].str.replace(',', '').astype(float)

        df['Transaction Description'] = df['Transaction Description'].str.replace(r'\d+,\d+\.\d+, ', '', regex=True)
        df['Transaction Description'] = df['Transaction Description'].str.replace(r', (\d{1,3}(?:,\d{3})*(?:\.\d{2}))$', '', regex=True)

        df = df[['Entry Date', 'Transaction Amount', 'Transaction Description', 'Statement Balance', 'Statement Balance 2']]
        df = df.rename(columns={'Transaction Amount': 'Transaction Type', 'Statement Balance': 'Transaction Amount', 'Statement Balance 2': 'Statement_Balance'})
        df.loc[df['Transaction Type'] == 'CASH WITHDRAWAL', 'Transaction Description'] = 'CASH WITHDRAWAL'
        df.loc[df['Transaction Type'] == 'DEBIT ADVICE', 'Transaction Description'] = 'Card Annual Fee'
        df.loc[df['Transaction Type'] == 'PROFIT PAID', 'Transaction Description'] = 'PROFIT PAID'

        df.loc[df['Transaction Type'] == 'INTEREST PAYMENT', 'Transaction Description'] = 'INTEREST PAYMENT'
        df.loc[df['Transaction Type'] == 'INT ON INT PAYMENT', 'Transaction Description'] = 'INT ON INT PAYMENT'

        df['flow'] = df['Transaction Amount'].apply(determine_flow)
        df['Transaction Amount'] = df['Transaction Amount'].str.replace('+', '', regex=False).str.replace('-', '', regex=False)
        df['Transaction Amount'] = df['Transaction Amount'].str.replace(',', '').astype(float)
        
        all_data.append(df)

    combined_df = pd.concat(all_data, ignore_index=True)
    excel_path = os.path.join(export_path, f"{excel_file_name}.csv")
    combined_df.to_csv(excel_path, index=False)
    print(f"Data exported to {excel_path}")
    
    messagebox.showinfo("Success", f"Data exported successfully to {excel_path}")

def process_files():
    folder_path = source_path_entry.get()
    export_path = export_path_entry.get()
    excel_file_name = excel_name_entry.get()
    if not folder_path or not excel_file_name or not export_path:
        messagebox.showerror("Error", "Folder path, export path, or Excel file name is missing")
        return
    
    pdf_files = [file for file in os.listdir(folder_path) if file.endswith('.pdf')]
    all_data = []

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()

        lines = text.split('\n')
        lines = remove_sections(lines, 'Maybank Islamic Berhad', 'Please notify us of any change of address in writing.')
        lines = remove_sections(lines, '15th Floor, Tower A, Dataran Maybank, 1, Jalan Maarof, 59000 Kuala Lumpur', '請通知本行在何地址更换。')
        lines = remove_sections(lines, 'ENTRY DATE', 'STATEMENT BALANCE')
        lines = remove_sections(lines, 'ENDING BALANCE :', 'TOTAL DEBIT :')


        filtered_lines = [line for line in lines if not any(s in line for s in strings_to_remove)]
        transactions = filtered_lines
        structured_data = []
        temp_entry = {}
        date_pattern = re.compile(r'\d{2}/\d{2}/\d{2}')

        for line in transactions:
            if date_pattern.match(line):
                if temp_entry:
                    structured_data.append(temp_entry)
                temp_entry = {"Entry Date": line, "Transaction Description": "", "Transaction Amount": "", "Statement Balance": ""}
            elif "Transaction Amount" in temp_entry and temp_entry["Transaction Amount"] and "Statement Balance" in temp_entry and temp_entry["Statement Balance"] == "":
                temp_entry["Statement Balance"] = line.strip()
            elif "Transaction Amount" in temp_entry and temp_entry["Transaction Amount"] == "":
                temp_entry["Transaction Amount"] = line.strip()
            else:
                if temp_entry:
                    temp_entry["Transaction Description"] += line.strip() + ", "

        if temp_entry:
            structured_data.append(temp_entry)

        for entry in structured_data:
            entry["Transaction Description"] = entry["Transaction Description"].rstrip(', ')

        df = pd.DataFrame(structured_data)
        df['Entry Date'] = pd.to_datetime(df['Entry Date'], format='%d/%m/%y', dayfirst=True).dt.date 
        df['Statement Balance 2'] = df['Transaction Description'].str.extract(r'(\d+,\d+\.\d+)')[0]
        df['Statement Balance 2'] = df['Statement Balance 2'].str.replace(',', '').astype(float)

        df['Transaction Description'] = df['Transaction Description'].str.replace(r'\d+,\d+\.\d+, ', '', regex=True)
        df['Transaction Description'] = df['Transaction Description'].str.replace(r', (\d{1,3}(?:,\d{3})*(?:\.\d{2}))$', '', regex=True)

        df = df[['Entry Date', 'Transaction Amount', 'Transaction Description', 'Statement Balance', 'Statement Balance 2']]
        df = df.rename(columns={'Transaction Amount': 'Transaction Type', 'Statement Balance': 'Transaction Amount', 'Statement Balance 2': 'Statement_Balance'})
        df.loc[df['Transaction Type'] == 'CASH WITHDRAWAL', 'Transaction Description'] = 'CASH WITHDRAWAL'
        df.loc[df['Transaction Type'] == 'DEBIT ADVICE', 'Transaction Description'] = 'Card Annual Fee'
        df.loc[df['Transaction Type'] == 'PROFIT PAID', 'Transaction Description'] = 'PROFIT PAID'
        df['flow'] = df['Transaction Amount'].apply(determine_flow)
        df['Transaction Amount'] = df['Transaction Amount'].str.replace('+', '', regex=False).str.replace('-', '', regex=False)
        df['Transaction Amount'] = df['Transaction Amount'].str.replace(',', '').astype(float)
        
        all_data.append(df)

    combined_df = pd.concat(all_data, ignore_index=True)
    excel_path = os.path.join(export_path, f"{excel_file_name}.csv")
    combined_df.to_csv(excel_path, index=False)
    print(f"Data exported to {excel_path}")
    
    messagebox.showinfo("Success", f"Data exported successfully to {excel_path}")

def remove_close_dates(data):
    valid_dates_indices = []
    i = 0
    while i < len(data):
        if re.match(r'\d{2}/\d{2}/\d{4}', data[i]):
            valid_dates_indices.append(i)
            i += 4
        else:
            i += 1
    filtered_data = [data[i] for i in range(len(data)) if i in valid_dates_indices or not re.match(r'\d{2}/\d{2}/\d{4}', data[i])]
    return filtered_data

def is_pure_number(s):
    # Remove spaces for the check
    s = s.replace(' ', '')
    # Check if the string is numeric and does not contain '.' or ','
    return s.isnumeric() and not any(c in s for c in ".,")
        
def process_CIMB_DEBIT_data():
    folder_path = source_path_entry.get()
    export_path = export_path_entry.get()
    excel_file_name = excel_name_entry.get()
    if not folder_path or not excel_file_name or not export_path:
        messagebox.showerror("Error", "Folder path, export path, or Excel file name is missing")
        return
    
    pdf_files = [file for file in os.listdir(folder_path) if file.endswith('.pdf')]
    all_data = []

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()

        lines = text.split('\n')
        lines = remove_sections(lines, 'Page / Halaman', 'ISLAMIC BBB-PPPP')

        filtered_lines = [line for line in lines if not any(s in line for s in strings_to_remove)]
        data = remove_close_dates(filtered_lines)
        data = [item for item in data if not is_pure_number(item)]
        data = [item if item != "99 SPEEDMART-2133" else "ninetynine speed mart" for item in data]

        final_structured_data = []

        i = 0
        while i < len(data):
            transaction = {}
            if data[i] == 'OPENING BALANCE':
                transaction['Date'] = '-'
                transaction['Transaction Type/Description'] = 'Opening Balance'
                i += 1
                transaction['Balance After Transaction'] = '-'
                transaction['Amount'] = data[i].strip()
                transaction['Beneficiary/Payee Name'] = '-'
                final_structured_data.append(transaction)
                i += 1
            elif re.match(r'\d{2}/\d{2}/\d{4}', data[i]):
                transaction['Date'] = data[i]
                i += 1

                description_lines = []
                while i < len(data) and not re.match(r'\d{2}/\d{2}/\d{4}', data[i]) and not re.match(r'^-?\d', data[i].strip()):
                    if data[i].strip():
                        description_lines.append(data[i].strip())
                    i += 1
            
                transaction['Transaction Type/Description'] = ', '.join(description_lines)

                if i < len(data) and re.match(r'^-?\d', data[i].strip()):
                    transaction['Amount'] = data[i].strip()
                    i += 1

                balance_line = data[i].strip() if i < len(data) else ""
                while not balance_line and i < len(data):
                    i += 1
                    balance_line = data[i].strip() if i < len(data) else ""
                transaction['Balance After Transaction'] = balance_line

                transaction['Beneficiary/Payee Name'] = '-'
                if description_lines:
                    transaction['Beneficiary/Payee Name'] = description_lines[0]

                final_structured_data.append(transaction)
            else:
                i += 1

        if final_structured_data:
            final_df = pd.DataFrame(final_structured_data)
            df = final_df
            df['Transaction Description2'] = df['Transaction Type/Description'].apply(lambda x: ' '.join(x.split()[1:]))
            df['Transaction Description'] = df['Transaction Description2'] + ', ' + df['Beneficiary/Payee Name']
            df.drop(columns=['Transaction Type/Description', 'Beneficiary/Payee Name'], inplace=True)
            for i in range(1, len(df)):
                if df.loc[i, 'Balance After Transaction'] > df.loc[i-1, 'Balance After Transaction']:
                    df.loc[i, 'output'] = 'deposit'
                else:
                    df.loc[i, 'output'] = 'withdrawal'
            all_data.append(df)

    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        combined_df['Transaction Description2'] = combined_df['Transaction Description2'].replace('Balance', 'Opening Balance')
        combined_df['Transaction Description'] = combined_df['Transaction Description2'].replace('Balance, -', 'Opening Balance')
        combined_df[['Date', 'Transaction Type']] = combined_df['Date'].str.extract(r'(\S+)\s(.*)')
        combined_df = combined_df[['Date', 'Transaction Type', 'Transaction Description', 'Transaction Description2','Amount', 'Balance After Transaction','output']]

        excel_path = os.path.join(export_path, f"{excel_file_name}.csv")
        combined_df.to_csv(excel_path, index=False)
        print(f"Data exported to {excel_path}")

        messagebox.showinfo("Success", f"Data exported successfully to {excel_path}")
    else:
        print("No Data to Export.")

# GUI code
root = tk.Tk()
root.title("MAE PDF File Processor")
root.configure(background='white')
root.geometry('800x200')

# Improved styling with ttk.Style
style = ttk.Style()
style.configure("TButton", font=('Arial', 10), background='lightgrey')
style.configure("TLabel", font=('Arial', 10), background='white')
style.configure("TEntry", font=('Arial', 10))

# Create a StringVar to hold the selection
processing_mode = tk.StringVar()
processing_mode.set("Maybank Debit Card Statement Processing")  # default value

# Create the dropdown menu
processing_mode_label = ttk.Label(root, text="Select Processing Mode:", background='white')
processing_mode_label.grid(row=3, column=0, sticky=tk.W, padx=(10, 5), pady=(5, 5))

processing_mode_dropdown = ttk.Combobox(root, textvariable=processing_mode)
processing_mode_dropdown['values'] = (
    "Maybank Debit Card Statement Processing",
    "Maybank Credit Card Statement Processing",
    "CIMB Debit Statement Processing",
    # "M2U Current Account Statement",
    "M2U Current Account Debit"
)
processing_mode_dropdown.grid(row=3, column=1, sticky=tk.EW, padx=(0, 10), pady=(5, 5))

# Directory selection for PDF files
source_path_entry = create_directory_selection_row(root, "Select Folder with PDFs:", select_directory, row=0)

# Directory selection for saving the Excel file
export_path_entry = create_directory_selection_row(root, "Select Export Path:", select_directory, row=1)

# Excel file name entry with improved layout and consistency, moved to after the export path
excel_name_label = ttk.Label(root, text="Enter Excel filename (without extension):", background='white')
excel_name_label.grid(row=2, column=0, sticky=tk.W, padx=(10, 5), pady=(5, 5))

excel_name_entry = ttk.Entry(root, width=60)
excel_name_entry.grid(row=2, column=1, columnspan=1, sticky=tk.EW, padx=(0, 10), pady=(5, 5))

# Process and export files button with consistent padding
style.configure("Green.TButton", font=('Arial', 10), background='lightgreen')
process_files_button = ttk.Button(
    root,
    text="Process Files and Export to Excel",
    command=selected_processing,
    style="Green.TButton"
)
process_files_button.grid(row=4, column=0, columnspan=3, padx=10, pady=(5, 10), sticky=tk.EW)

root.grid_columnconfigure(1, weight=1)  # Make the second column expandable

root.mainloop()