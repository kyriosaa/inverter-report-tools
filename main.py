import os
import io
import pandas as pd
import win32com.client
from datetime import datetime, timedelta

# if ur downloading from github pls make a "private.py" folder and add the folder path to the variables
# MASTER_CSV_PATH = r"C:\path\to\folder\Master_Inverter_Data.csv"
# etc etc
import private
from config.colors import BOLD, RED, CYAN, YELLOW, GREEN, RESET

# change ts to the email subject to pull the report from
EMAIL_SUBJECT = "Taiwan Solar Floating - Monthly-Inverter Report"


if not os.path.exists(private.TEMP_DOWNLOAD_FOLDER):
    os.makedirs(private.TEMP_DOWNLOAD_FOLDER)

def get_last_updated_date():
    # checks the master CSV to find the last recorded date
    if os.path.exists(private.MASTER_CSV_PATH):
        df = pd.read_csv(private.MASTER_CSV_PATH)
        if not df.empty:
            # get the max date and convert to a datetime object
            latest_date_str = df['Date'].max()
            return datetime.strptime(latest_date_str, '%Y-%m-%d').date()
    
    # default to looking 31 days back if no master file
    return (datetime.today() - timedelta(days=31)).date()

# for if u have multiple outlook accounts
def select_outlook_inbox(outlook):
    print(f"\n{BOLD}{CYAN}Available Outlook Accounts:{RESET}")
    stores = []
    
    for i in range(1, outlook.Stores.Count + 1): # note that outlook indexes start at one instead of 0
        store = outlook.Stores.Item(i)
        stores.append(store)
        print(f"{i}. {store.DisplayName}")
        
    while True:
        try:
            choice = int(input(f"\n{BOLD}{CYAN}Select account (1-{stores.Count}): {RESET}")).strip()
            if 1 <= choice <= len(stores):
                selected_store = stores[choice - 1]
                # 6 is the default inbox folder
                return selected_store.GetDefaultFolder(6)
            else:
                print(f"{BOLD}{RED}[ERROR]{RESET} Please select a valid number from the list.")
        except ValueError:
            print(f"{BOLD}{RED}[ERROR]{RESET} Please enter a number.")
        except Exception as err:
            print(f"{BOLD}{RED}[ERROR]{RESET} Error accessing folder: {err}. Please try again.")

# scans local outlook app for emails newer than the last updated date and downloads attachments
def fetch_missing_reports_from_outlook(last_date):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # select inbox
    inbox = select_outlook_inbox(outlook)
    print(f"\n{BOLD}{YELLOW}[STATUS]{RESET} Scanning Inbox for '{inbox.Parent.Name}' for reports after {last_date}...")
    
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True) # sort newest first
    
    downloaded_files = []
    
    for message in messages:
        try:
            # get the date from the email's arrival time bcs the automated inverter reports have the same name
            received_date = message.ReceivedTime.date()
            
            if received_date <= last_date:
                break
                
            # check subject matching
            if EMAIL_SUBJECT.lower() in message.Subject.lower():
                for attachment in message.Attachments:
                    if attachment.FileName.endswith('.csv') or attachment.FileName.endswith('.xlsx'):
                        file_path = os.path.join(private.TEMP_DOWNLOAD_FOLDER, f"{received_date}_{attachment.FileName}")
                        attachment.SaveAsFile(file_path)
                        downloaded_files.append((received_date, file_path))
                        print(f"{BOLD}{YELLOW}[STATUS]{RESET} Downloaded report for {received_date}")
        except Exception:
            continue
            
    return downloaded_files

# reads the report ahd automatically handles both .xlsx and .csv files
def safe_read_report(file_path):
    # if its excel we can just let Pandas handle it directly
    if file_path.endswith('.xlsx'):
        temp_df = pd.read_excel(file_path, header=None, nrows=15)
        header_idx = 0
        for i, row in temp_df.iterrows():
            # convert the row to a string to check for columns
            row_str = " ".join([str(x) for x in row.values])
            if 'Plant Name' in row_str and 'Yield' in row_str:
                header_idx = i
                break
        return pd.read_excel(file_path, header=header_idx)
        
    else:
        # if its a CSV file, use byte-decoder
        with open(file_path, 'rb') as f:
            raw_data = f.read()
            
        encodings_to_try = ['utf-8-sig', 'utf-8', 'gbk', 'big5', 'utf-16', 'cp1252', 'latin1']
        decoded_text = None
        
        for enc in encodings_to_try:
            try:
                text = raw_data.decode(enc)
                if '\x00' in text and not enc.startswith('utf-16'):
                    continue
                decoded_text = text
                break 
            except UnicodeDecodeError:
                continue
                
        if decoded_text is None:
            decoded_text = raw_data.decode('utf-8', errors='replace')
            
        lines = decoded_text.splitlines()
        header_idx = 0
        for i, line in enumerate(lines[:15]): 
            if 'Plant Name' in line and 'Yield' in line:
                header_idx = i
                break
                
        return pd.read_csv(io.StringIO(decoded_text), header=header_idx, engine='python', on_bad_lines='skip')

# reads files, cleans data, then appends to master CSV
def process_and_append_data(downloaded_files):
    if not downloaded_files:
        print(f"{BOLD}{GREEN}[DONE]{RESET} No new reports found to process.")
        return False

    all_new_data = []
    
    for email_date, file_path in downloaded_files:
        df = safe_read_report(file_path)
        
        # strip any whitespace from column names (stuff like "Yield (kWh) ")
        df.columns = df.columns.astype(str).str.strip()
        
        try:
            # filter only the columns we need for the database
            clean_df = df[['Plant Name', 'Device Name', 'Yield (kWh)']].copy()
        except KeyError:
            print(f"{BOLD}{RED}[ERROR]{RESET} Could not find columns in report for {email_date}. Found columns: {df.columns.tolist()}")
            os.remove(file_path)
            continue

        # add the date from the email
        clean_df['Date'] = email_date.strftime('%Y-%m-%d')
        all_new_data.append(clean_df)
        
        # delete temp file after we're done
        os.remove(file_path)

    if not all_new_data:
        print(f"{BOLD}{RED}[ERROR]{RESET} No usable data was extracted from the downloaded files.")
        return False

    # combine all new days into one table
    new_data_df = pd.concat(all_new_data, ignore_index=True)
    
    # append to master CSV
    if os.path.exists(private.MASTER_CSV_PATH):
        new_data_df.to_csv(private.MASTER_CSV_PATH, mode='a', header=False, index=False)
    else:
        new_data_df.to_csv(private.MASTER_CSV_PATH, mode='w', header=True, index=False)
        
    print(f"{BOLD}{GREEN}[DONE]{RESET} Successfully appended {len(new_data_df)} records to the Master CSV.")
    return True

# this is the view that makes it easy to read
def generate_report_view():
    if not os.path.exists(private.MASTER_CSV_PATH):
        return
        
    print(f"{BOLD}{YELLOW}[STATUS]{RESET} Generating updated visual report...")
    df = pd.read_csv(private.MASTER_CSV_PATH)
    
    df_clean = df.drop_duplicates(subset=['Plant Name', 'Device Name', 'Date'])
    pivot_df = df_clean.pivot(index=['Plant Name', 'Device Name'], columns='Date', values='Yield (kWh)')
    
    # sort by the actual inverter number
    pivot_df = pivot_df.reset_index()
    pivot_df['Inverter_Number'] = pivot_df['Device Name'].str.extract(r'(\d+)[^\d]*$').astype(float)
    
    # sort hierarchically
    # first by plant, then by the extracted number, then by the original name as a fallback
    pivot_df = pivot_df.sort_values(by=['Plant Name', 'Inverter_Number', 'Device Name'])
    pivot_df = pivot_df.drop(columns=['Inverter_Number']).set_index(['Plant Name', 'Device Name'])
    
    pivot_df.to_csv(private.REPORT_VIEW_PATH)
    print(f"{BOLD}{GREEN}[DONE]{RESET} Visual report saved to: {private.REPORT_VIEW_PATH}")

if __name__ == "__main__":
    last_date = get_last_updated_date()
    files_to_process = fetch_missing_reports_from_outlook(last_date)
    
    data_added = process_and_append_data(files_to_process)
    
    if data_added:
        generate_report_view()
    
    print(f"{BOLD}{GREEN}Complete.{RESET}")