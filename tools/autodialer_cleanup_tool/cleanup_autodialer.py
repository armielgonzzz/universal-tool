import os
import re
import pandas as pd
import warnings
import dropbox
from dotenv import load_dotenv
from sqlalchemy import create_engine
from urllib.parse import quote

warnings.simplefilter(action='ignore', category=pd.errors.SettingWithCopyWarning)

disposition_query = """
SELECT
	dnis_to
FROM
	max_outbound_calls
WHERE
	primary_disposition IS NOT NULL AND
    primary_disposition IN 
	(
		'Business/ Work number',
		'Sold Interests',
		'Incorrect contact / Wrong number',
		'Do Not Call Again (remove from list)',
		'Invalid Number',
		'Proactive Identified - Answering Machine Left Message',
		'Answering Machine Left Message'
	)
GROUP BY
	dnis_to
"""

six_months_query = """
WITH ranked_calls AS (
    SELECT
        dnis_to,
        ROW_NUMBER() OVER (PARTITION BY dnis_to ORDER BY start_time DESC) AS date_rank
    FROM
        max_outbound_calls
    WHERE
        primary_disposition IN ('Lead Not interested', 'Uncooperative Lead')
        AND start_time BETWEEN DATE_ADD(CURRENT_DATE, INTERVAL -6 MONTH) AND CURRENT_DATE
)
SELECT
    dnis_to
FROM
    ranked_calls
WHERE
    date_rank = 1;
"""

def extract_list_cleaner_file(auth_code: str, local_path: str, dropbox_path: str):
    dbx = dropbox.Dropbox(auth_code)
    metadata, response = dbx.files_download(dropbox_path)
    with open(local_path, 'wb') as f:
        f.write(response.content)

def read_cm_live_db() -> 'tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame | None]':

    try:
        # Create database engine
        host = os.getenv('DB_HOST')
        user = os.getenv('DB_USER')
        name = os.getenv('DB_NAME')
        password = os.getenv('DB_PASSWORD')
        engine = create_engine(f'mysql+pymysql://{user}:{quote(password)}@{host}/{name}')

        print(f'Reading Community Minerals Database')

        disposition_df = pd.read_sql_query(disposition_query, engine)
        six_months_df = pd.read_sql_query(six_months_query, engine)

        disposition_df['dnis_to'] = pd.to_numeric(disposition_df['dnis_to'], errors='coerce')
        disposition_df['dnis_to'] = disposition_df['dnis_to'].astype('Int64')
        disposition_set = set(disposition_df.loc[disposition_df['dnis_to'].notna(), 'dnis_to'].map(int))

        six_months_df['dnis_to'] = pd.to_numeric(six_months_df['dnis_to'], errors='coerce')
        six_months_df['dnis_to'] = six_months_df['dnis_to'].astype('Int64')
        months_set = set(six_months_df.loc[six_months_df['dnis_to'].notna(), 'dnis_to'].map(int))

        return disposition_set, months_set

    except Exception as e:
        raise RuntimeError(f"An error occurred while reading from the database: {e}")

    finally:
        engine.dispose()


def export_output(df: pd.DataFrame, file_path: str, save_path: str) -> None:

    # Capitalization of all names
    columns_to_transform = ['Owner','Combined Name', 'First Name', 'Middle Name', 'Last Name']
    df[columns_to_transform] = df[columns_to_transform].applymap(lambda x: x.title() if isinstance(x, str) else x)

    filename = os.path.basename(file_path)
    if filename.endswith('.csv'):
        df.to_csv(f"{save_path}/(Clean file) {filename}", index=False)
    
    elif filename.endswith('.xlsx'):
        df.to_excel(f"{save_path}/(Clean file) {filename}", index=False)

    elif filename.endswith('.xlsb'):
        df.to_excel(f"{save_path}/(Clean file) {filename.split('.')[0]}.xlsx", index=False)
    
    else:
        print("No output generated. Invalid file format")

def read_file(path: str):

    if path.endswith('.csv'):
        return pd.read_csv(path, low_memory=False)

    elif path.endswith('.xlsx'):
        return pd.read_excel(path)
    
    elif path.endswith('.xlsb'):
        return pd.read_excel(path, engine='pyxlsb', sheet_name=['List'])['List']
    
    else:
        raise ValueError("Invalid file format: Please provide a .csv, .xlsx or .xlsb file.")
    
def is_valid_phone(phone):
    return bool(re.fullmatch(r'\d{10,15}', phone))

def remove_phone_dupes(df: pd.DataFrame) -> pd.DataFrame:

    df['original_index'] = df.index
    phone_columns = ['phone1', 'phone2', 'phone3', 'phone4', 'phone5']
    sorted_df = df.sort_values(by=phone_columns, ascending=False)

    for phone in phone_columns:
        sorted_df.loc[(sorted_df[phone].notna()) & (sorted_df[phone].duplicated(keep='last')), phone] = pd.NA
    
    phone_filter_mask = (sorted_df['phone5'].isin(sorted_df['phone2'])) \
    | (sorted_df['phone5'].isin(sorted_df['phone3'])) \
    | (sorted_df['phone5'].isin(sorted_df['phone4'])) \
    | (sorted_df['phone5'].isin(sorted_df['phone1']))
    sorted_df.loc[(sorted_df['phone5'].notna()) & (phone_filter_mask), 'phone5'] = pd.NA

    phone_filter_mask = (sorted_df['phone4'].isin(sorted_df['phone3'])) \
    | (sorted_df['phone4'].isin(sorted_df['phone2'])) \
    | (sorted_df['phone4'].isin(sorted_df['phone1']))
    sorted_df.loc[(sorted_df['phone4'].notna()) & (phone_filter_mask), 'phone4'] = pd.NA

    phone_filter_mask = (sorted_df['phone3'].isin(sorted_df['phone2'])) \
    | (sorted_df['phone3'].isin(sorted_df['phone1']))
    sorted_df.loc[(sorted_df['phone3'].notna()) & (phone_filter_mask), 'phone3'] = pd.NA

    phone_filter_mask = (sorted_df['phone2'].isin(sorted_df['phone1']))
    sorted_df.loc[(sorted_df['phone2'].notna()) & (phone_filter_mask), 'phone2'] = pd.NA
            
    sorted_df = sorted_df.sort_values(by='original_index')
    sorted_df = sorted_df.drop(columns='original_index')
    return sorted_df

def clean_contact_id_deal_id(df: pd.DataFrame, id_set: set) -> pd.DataFrame:

    if 'contact_id' in df.columns:
        df['contact_id'] = df['contact_id'].apply(pd.to_numeric, errors='coerce').astype('Int64')
        df = df[~df['contact_id'].isin(id_set)]
    
    if 'Deal ID' in df.columns:
        df = df[df['Deal ID'].isna()]

    final_df = df[~df[['phone1', 'phone2', 'phone3', 'phone4', 'phone5']].isna().all(axis=1)]

    return final_df

def get_phone_set(cleaner_file: str) -> set:

    list_cleaner_df = pd.read_excel(cleaner_file,
                                    sheet_name=['CCM+CH+MVPC+MVPT+JC+RC+PD',
                                                'DNC',
                                                'CallOut-14d+TextOut-30d',
                                                'PDConvDup',
                                                'FromOtherList'],
                                    header=None)
    
    final_list_cleaner_df = pd.concat([list_cleaner_df['CCM+CH+MVPC+MVPT+JC+RC+PD'],
                                        list_cleaner_df['DNC'],
                                        list_cleaner_df['CallOut-14d+TextOut-30d'],
                                        list_cleaner_df['PDConvDup'],
                                        list_cleaner_df['FromOtherList']])

    # Clean up the list and filter for valid phone numbers
    valid_phone_set = set(int(phone) for phone in map(str, final_list_cleaner_df[0].tolist()) if is_valid_phone(phone))
    return valid_phone_set

def get_id_set(cleaner_file: str) -> set:

    unique_db_df_list = pd.read_excel(cleaner_file,
                                      sheet_name=['UniqueDB ID'])
    unique_db_df = unique_db_df_list['UniqueDB ID']
    valid_numbers = pd.to_numeric(unique_db_df['Deal - Unique Database ID'], errors='coerce')
    valid_numbers = valid_numbers.dropna().astype(int)
    valid_id_set = set(valid_numbers)
    return valid_id_set

def main(auth_code: str, list_files: tuple, save_path: str):

    try:
        
        print("Preparing List Cleaner")
        local_list_cleaner_path = './data/List Cleaner.xlsx'
        dropbox_list_cleaner_path = '/List Cleaner & JC DNC/New List Cleaner.xlsx'

        extract_list_cleaner_file(auth_code, local_list_cleaner_path, dropbox_list_cleaner_path)

        valid_phone_set = get_phone_set(local_list_cleaner_path)
        valid_id_set = get_id_set(local_list_cleaner_path)
        disposition_set, months_set = read_cm_live_db()
        
        for list_file in list_files:

            print(f"Processing file {os.path.basename(list_file)}")
            list_df = read_file(list_file)

            # Convert phone numbers to int
            phone_columns = ['phone1', 'phone2', 'phone3', 'phone4', 'phone5']
            list_df[phone_columns] = (
                list_df[phone_columns]
                .apply(pd.to_numeric, errors='coerce')
                .astype('Int64')
            )

            # Search phones in cleaner file
            output_df = list_df[~list_df[phone_columns].isin(valid_phone_set).any(axis=1)]

            # Remove Company contact type
            column_name = next((col for col in output_df.columns if col.strip().lower() == 'contact_type'), None)
            if column_name:
                output_df = output_df[output_df[column_name].str.lower() != 'company']
            
            # Remove duplicates
            removed_dupes_df = remove_phone_dupes(output_df)

            # Clean df based on contact id and deal id
            clean_contact_deal_df = clean_contact_id_deal_id(removed_dupes_df, valid_id_set)

            # Check if has existing dispositions
            clean_dispo_df = clean_contact_deal_df[~clean_contact_deal_df[phone_columns].isin(disposition_set).any(axis=1)]

            # Check if within 6 months for specific dispositions
            final_df = clean_dispo_df[~clean_dispo_df[phone_columns].isin(months_set).any(axis=1)]

            # Export to save path
            export_output(final_df, list_file, save_path)
        
        print("Sucessfully processed all files")

    except Exception as e:
        print(f"An error occurred: {e}")
        raise RuntimeError   

if __name__ == "__main__":
    main()
