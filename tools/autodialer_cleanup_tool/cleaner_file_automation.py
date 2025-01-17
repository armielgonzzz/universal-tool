import re
import os
import pandas as pd
import warnings
import dropbox
import datetime
import webbrowser
from datetime import timedelta
from dotenv import load_dotenv
from io import BytesIO
from openpyxl import load_workbook

warnings.simplefilter(action='ignore', category=pd.errors.SettingWithCopyWarning)
load_dotenv(dotenv_path='misc/.env')
APP_KEY = os.getenv('DROPBOX_APP_KEY')
APP_SECRET = os.getenv('DROPBOX_APP_SECRET')

def dropbox_authentication() -> str:
    auth_flow = dropbox.DropboxOAuth2FlowNoRedirect(APP_KEY, APP_SECRET)
    authorize_url = auth_flow.start()

    if authorize_url:
        webbrowser.open(authorize_url)

def export_to_dropbox(list_cleaner_file_path, dbx) -> None:
    datetime_now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    list_cleaner_dropbox_path = f'/list cleaner & jc dnc/list_cleaner_file/{datetime_now} - List Cleaner File.xlsx'
    with open(list_cleaner_file_path, 'rb') as f:
        print("Uploading to List Cleaner File to Dropbox")
        dbx.files_upload(f.read(), list_cleaner_dropbox_path, mode=dropbox.files.WriteMode.overwrite)
    
    print("Sucessfully uploaded Updated List Cleaner File to Dropbox")

def read_dropbox_file(path: str, dbx):
    metadata, response = dbx.files_download(path)
    if path.endswith('.csv'):
        return pd.read_csv(BytesIO(response.content), low_memory=False)
    elif path.endswith('.xlsx'):
        return pd.read_excel(BytesIO(response.content))
    else:
        raise ValueError("Invalid file format: Please provide a .csv, .xlsx or .xlsb file.")
    
def add_pd_phones(path: str, excel_file: str, dbx):
    print("Adding Sheet ContMgt+MVP+JC+PD+RC")
    # Drop columns that is not needed
    df = read_dropbox_file(path, dbx)
    df.drop(columns=['Deal - ID', 'Deal - Last RVM Date', 'Deal - RVM Dates'],
            axis=1,
            inplace=True)

    # Select only the phone columns (Phone 1 to Phone 10)
    phone_columns = [f'Person - Phone {i}' for i in range(1, 11)]
    phones_df = df[phone_columns]

    # Initialize a list to hold all phone numbers
    all_phone_numbers = []

    # Iterate over each column and each row to extract phone numbers
    for column in phone_columns:
        phones_df[column] = phones_df[column].fillna('')  # Replace NaN with an empty string
        for phone_entry in phones_df[column]:
            # Convert each entry to a string, handle floats (remove .0), and split comma-separated values
            phone_entry = str(phone_entry).replace('.0', '')  # Remove `.0` from floats
            valid_phone_pattern = re.compile(r'^\+?\d{10}$')  # A simple pattern for phone numbers (you can adjust it)
            all_phone_numbers.extend([phone.strip() for phone in phone_entry.split(',') if phone.strip() and valid_phone_pattern.match(phone.strip())])
            # all_phone_numbers.extend([phone.strip() for phone in phone_entry.split(',') if phone.strip()])

    # Create a DataFrame from the extracted phone numbers
    result_df = pd.DataFrame({'Phone Number': all_phone_numbers})

    result_df['Phone Number'] = result_df['Phone Number'] \
        .str.replace("(", "") \
        .str.replace(")", "") \
        .str.replace("-", "") \
        .str.replace(" ", "")
    
    result_df.drop_duplicates(subset=['Phone Number'], inplace=True)
    
    # Use ExcelWriter to append the DataFrame to the existing file
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        
        result_df.to_excel(writer, sheet_name='ContMgt+MVP+JC+PD+RC', index=False)

def add_unique_db(path: str, excel_file: str, dbx):
    print("Adding Sheet UniqueDB ID")
    df = read_dropbox_file(path, dbx)
    # Now handle the "Deal - Unique Database ID" column similarly
    deal_column = 'Deal - Unique Database ID'

    # Ensure "Deal - Unique Database ID" has NaN values replaced with empty strings
    df[deal_column] = df[deal_column].fillna('')

    # Initialize a list to hold all database IDs
    all_deals = []

    # Iterate over each row in the "Deal - Unique Database ID" column
    for deal_entry in df[deal_column]:
        # Split by " | " and add each entry to the list
        all_deals.extend([deal.strip() for deal in deal_entry.split('|') if deal.strip()])

    # Create a DataFrame for the Deal IDs
    deal_df = pd.DataFrame({'Deal - Unique Database ID': all_deals})

    deal_df.drop_duplicates(subset=['Deal - Unique Database ID'], inplace=True)

    # Use ExcelWriter to append the DataFrame to the existing file
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        
        deal_df.to_excel(writer, sheet_name='UniqueDB ID', index=False)

def add_remove_list(path: str, excel_file: str, dbx):
    print("Updating Sheet ContMgt+MVP+JC+PD+RC")
    df = read_dropbox_file(path, dbx)
    phone_columns = [f'Person - Phone {i}' for i in range(2, 11)]
    phone_columns.extend(['Person - Phone - Work', 'Person - Phone - Home', 'Person - Phone - Mobile', 'Person - Phone - Other'])
    phones_df = df[phone_columns]

    all_phone_numbers = []

    # Iterate over each column and each row to extract phone numbers
    for column in phone_columns:
        phones_df[column] = phones_df[column].fillna('')  # Replace NaN with an empty string
        for phone_entry in phones_df[column]:
            # Convert each entry to a string, handle floats (remove .0), and split comma-separated values
            phone_entry = str(phone_entry).replace('.0', '')  # Remove `.0` from floats
            valid_phone_pattern = re.compile(r'^\+?\d{10}$')  # A simple pattern for phone numbers (you can adjust it)
            all_phone_numbers.extend([phone.strip() for phone in phone_entry.split(',') if phone.strip() and valid_phone_pattern.match(phone.strip())])
            # all_phone_numbers.extend([phone.strip() for phone in phone_entry.split(',') if phone.strip()])

    # Create a DataFrame from the extracted phone numbers
    result_df = pd.DataFrame({'Phone Number': all_phone_numbers})

    result_df['Phone Number'] = result_df['Phone Number'] \
        .str.replace("(", "") \
        .str.replace(")", "") \
        .str.replace("-", "") \
        .str.replace(" ", "")
    
    result_df.drop_duplicates(subset=['Phone Number'], inplace=True)

    book = load_workbook(excel_file)

    # Check if the sheet exists and load it into a DataFrame
    if 'ContMgt+MVP+JC+PD+RC' in book.sheetnames:
        # Read the existing sheet into a DataFrame
        existing_df = pd.read_excel(excel_file, sheet_name='ContMgt+MVP+JC+PD+RC')
        
        # Append the new rows to the existing DataFrame
        updated_df = pd.concat([existing_df, result_df], ignore_index=True)
    else:
        # If the sheet doesn't exist, just use the new DataFrame
        updated_df = result_df

    # Write the updated DataFrame back to the sheet
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        
        updated_df.to_excel(writer, sheet_name='ContMgt+MVP+JC+PD+RC', index=False, header=False)

def add_jc(path: str, excel_file: str, dbx):
    print("Adding Sheet JCSMS-Received")
    metadata, response = dbx.files_download(path)
    df = pd.read_excel(BytesIO(response.content),
                       sheet_name="Messages Details",
                       header=6,
                       usecols=['Client Number', 'Delivery Status'])

    df['Client Number'] = df['Client Number'].astype('Int64')
    result_df = df[df['Client Number'].notna()].copy()  # Ensure we work on a copy of the filtered DataFrame
    result_df['Client Number'] = result_df['Client Number'].astype('Int64').astype(str).str[1:]

    received_df = result_df[result_df['Delivery Status'].str.lower() == 'received'][['Client Number']]
    sent_df = result_df[(result_df['Delivery Status'].str.lower() == 'sent') | (result_df['Delivery Status'].str.lower() == 'delivered')][['Client Number']]

    received_df.drop_duplicates(subset=['Client Number'], inplace=True)
    sent_df.drop_duplicates(subset=['Client Number'], inplace=True)

    # Use ExcelWriter to append the DataFrame to the existing file
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        
        received_df.to_excel(writer, sheet_name='JCSMS-Received', index=False, header=False)

    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        print("Adding Sheet JCSMS-Sent")
        sent_df.to_excel(writer, sheet_name='JCSMS-Sent', index=False, header=False)

def add_sly(path: str, excel_file: str, dbx):
    print("Adding Sheet DNC")
    df = read_dropbox_file(path, dbx)
    df.drop_duplicates(inplace=True)
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        
        df.to_excel(writer, sheet_name='DNC', index=False, header=False)

def add_contact_center(df: pd.DataFrame, excel_file: str, dbx):
    print("Adding Sheet ContactMgtLogs")
    fourteen_days_ago = pd.to_datetime('today') - timedelta(days=14)
    df['Date'] = pd.to_datetime(df['Date'])
    df = df[df['Media Type Name'] != 'E-Mail']
    inbound_df = df[df['Skill Direction'] == 'Inbound'][['ANI/From']]
    outbound_df = df[(df['Skill Direction'] == 'Outbound') & (df['Date'] >= fourteen_days_ago)][['DNIS/To']]

    inbound_df.drop_duplicates(inplace=True)
    outbound_df.drop_duplicates(inplace=True)

    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        inbound_df.to_excel(writer, sheet_name='ContactMgtLogs', index=False, header=False)

    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        print("Adding Sheet Outbound-2weeks")
        outbound_df.to_excel(writer, sheet_name='Outbound-2weeks', index=False, header=False)

def add_mvp(path: str, excel_file: str, dbx):
    print("Adding Sheet MVPLogs")
    df = read_dropbox_file(path, dbx)
    df['Sender Number'] = df['Sender Number'].astype('Int64')
    valid_number_df = df[df['Sender Number'].notna()].astype(str)
    valid_number_df['Sender Number'] = valid_number_df['Sender Number'].str[1:]
    mvp_df = valid_number_df[valid_number_df['Direction'] == 'Inbound']
    final_df = mvp_df[mvp_df['Sender Number'].str.len() == 10][['Sender Number']]
    final_df.drop_duplicates(subset=['Sender Number'], inplace=True)
    
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        final_df.to_excel(writer, sheet_name='MVPLogs', index=False, header=False)

def add_rc(df: pd.DataFrame, excel_file: str, dbx):
    print("Adding Sheet RCSMS-Received")
    thirty_days_ago = pd.to_datetime('today', utc=True) - timedelta(days=30)
    df['Creation Time (UTC)'] = pd.to_datetime(df['Creation Time (UTC)'], utc=True)
    received_df = df[df['Direction'] == 'Inbound'][['From']]
    sent_df = df[(df['Direction'] == 'Outbound') & (df['Creation Time (UTC)'] >= thirty_days_ago)][['To']]
    received_df['From'] = received_df['From'].str.findall(r'\d+').apply(lambda x: ''.join(x)).str[1:]
    sent_df['To'] = sent_df['To'].str.findall(r'\d+').apply(lambda x: ''.join(x)).str[1:]

    received_df.drop_duplicates(subset=['From'], inplace=True)
    sent_df.drop_duplicates(subset=['To'], inplace=True)

    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        received_df.to_excel(writer, sheet_name='RCSMS-Received', index=False, header=False)

    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        print("Adding Sheet RCSMS-Sent")
        sent_df.to_excel(writer, sheet_name='RCSMS-Sent', index=False, header=False)

    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        pd.DataFrame().to_excel(writer, sheet_name='FromOtherList', index=False, header=False)

def get_latest_file(folder_path, dbx):
    try:
        # List files in the folder
        response = dbx.files_list_folder(folder_path)

        latest_file = None
        latest_timestamp = None

        for entry in response.entries:
            if isinstance(entry, dropbox.files.FileMetadata):
                # Compare the 'client_modified' timestamp to find the latest file
                if latest_timestamp is None or entry.client_modified > latest_timestamp:
                    latest_file = entry
                    latest_timestamp = entry.client_modified

        return latest_file.path_lower if latest_file else None

    except Exception as e:
        print(f"Error getting latest file from {folder_path}: {e}")
        return None

def load_sheets(path, dbx, conversion_dict):
    try:
        # List files and folders in the current path
        response = dbx.files_list_folder(path)
        for entry in response.entries:
            if isinstance(entry, dropbox.files.FileMetadata):
                file_path = entry.path_lower
                folder_path = "/".join(file_path.split('/')[:-1])
                folder_name = file_path.split('/')[:-1][-1]

                # Check if valid folder name then get corresponding function per folder
                if folder_name in conversion_dict:
                    file_type_function = conversion_dict[folder_name]

                    # Process the latest only
                    if file_path == get_latest_file(folder_path, dbx):
                        file_type_function(file_path, './data/List Cleaner File.xlsx', dbx)
                        
            elif isinstance(entry, dropbox.files.FolderMetadata):
                load_sheets(entry.path_lower, dbx, conversion_dict)

        # Handle pagination
        while response.has_more:
            response = dbx.files_list_folder_continue(response.cursor)
            for entry in response.entries:
                if isinstance(entry, dropbox.files.FileMetadata):
                    file_path = entry.path_lower
                    folder_path = "/".join(file_path.split('/')[:-1])
                    folder_name = file_path.split('/')[:-1][-1]

                    # Check if valid folder name then get corresponding function per folder
                    if folder_name in conversion_dict:
                        file_type_function = conversion_dict[folder_name]

                        # Process the latest only
                        if file_path == get_latest_file(folder_path, dbx):
                            file_type_function(file_path, './data/List Cleaner File.xlsx', dbx)
                            
                elif isinstance(entry, dropbox.files.FolderMetadata):
                    load_sheets(entry.path_lower, dbx, conversion_dict)

    except dropbox.exceptions.ApiError as e:
        print(f"Error accessing path '{path}': {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

def concat_contact_center_files(path: str, dbx):

    result = dbx.files_list_folder(path)
    contact_center_files = result.entries
    df_list = []

    for file in contact_center_files:
        if isinstance(file, dropbox.files.FileMetadata):
            file_path = file.path_lower
            df = read_dropbox_file(file_path, dbx)
            df_list.append(df)
    
    if df_list:
        combined_df = pd.concat(df_list)
        return combined_df
    
def concat_rc_files(path: str, dbx):

    result = dbx.files_list_folder(path)
    rc_files = result.entries
    df_list = []

    for file in rc_files:
        if isinstance(file, dropbox.files.FileMetadata):
            file_path = file.path_lower
            df = read_dropbox_file(file_path, dbx)
            df_list.append(df)
    
    if df_list:
        combined_df = pd.concat(df_list)
        return combined_df


def create_local_list_cleaner(path: str) -> None:

    print("Creating new local list cleaner file")

    # Check if the file exists and remove it
    if os.path.exists(path):
        os.remove(path)

    df = pd.DataFrame()

    # Save the empty DataFrame to an Excel file
    df.to_excel(path, index=False)

def check_user_folder_paths(dbx: dropbox.Dropbox):
    result = dbx.files_list_folder(path="", recursive=True)
    for entry in result.entries:
        if isinstance(entry, dropbox.files.FolderMetadata):
            print(entry.path_display)
            if entry.path_display == '/List Cleaner & JC DNC':
                return entry.path_display
                

def main(auth_code: str):

    try:
        auth_flow = dropbox.DropboxOAuth2FlowNoRedirect(APP_KEY, APP_SECRET)   
        oauth_result = auth_flow.finish(auth_code)
        dbx = dropbox.Dropbox(oauth_result.access_token)
        
        # Create constant variables
        local_list_cleaner_path = './data/List Cleaner File.xlsx'
        conversion_dict = {
            "jc": add_jc,
            "mvp": add_mvp,
            "pd_db": add_unique_db,
            "pd_phone": add_pd_phones,
            "pd_remove": add_remove_list,
            "sly": add_sly
        }

        path = check_user_folder_paths(dbx)

        # Remove and create new local list cleaner file
        create_local_list_cleaner(local_list_cleaner_path)

        # Update the list cleaner file
        load_sheets(path, dbx, conversion_dict)

        # Process contact_center folder files
        contact_center_df = concat_contact_center_files(f'{path}/contact_center', dbx)
        add_contact_center(contact_center_df, './data/List Cleaner File.xlsx', dbx)

        # Process rc folder files
        rc_df = concat_rc_files(f'{path}/rc', dbx)
        add_rc(rc_df, './data/List Cleaner File.xlsx', dbx)

        # Upload to dropbox
        export_to_dropbox(local_list_cleaner_path, dbx)
    
    except Exception as e:
        print(f"An error occured: {e}")
        raise RuntimeError

if __name__ == "__main__":
    main()
