import os
import re
import pandas as pd
import warnings

warnings.simplefilter(action='ignore', category=pd.errors.SettingWithCopyWarning)

def export_output(df: pd.DataFrame, file_path: str, save_path: str) -> None:

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
                                    sheet_name=['ContMgt+MVP+JC+PD+RC',
                                                'DNC',
                                                'SMS-Sent',
                                                'Outbound-2weeks',
                                                'FromOtherList'],
                                    header=None)
    
    final_list_cleaner_df = pd.concat([list_cleaner_df['ContMgt+MVP+JC+PD+RC'],
                                        list_cleaner_df['DNC'],
                                        list_cleaner_df['SMS-Sent'],
                                        list_cleaner_df['Outbound-2weeks'],
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

def main(cleaner_file: str, list_files: tuple, save_path: str):

    try:
        
        print("Preparing List Cleaner")

        valid_phone_set = get_phone_set(cleaner_file)
        valid_id_set = get_id_set(cleaner_file)
        
        for list_file in list_files:

            print(f"Processing file {os.path.basename(list_file)}")
            list_df = read_file(list_file)

            # Convert phone numbers to int
            list_df[['phone1', 'phone2', 'phone3', 'phone4', 'phone5']] = (
                list_df[['phone1', 'phone2', 'phone3', 'phone4', 'phone5']]
                .apply(pd.to_numeric, errors='coerce')
                .astype('Int64')
            )

            # Search phones in cleaner file
            output_df = list_df[~list_df[['phone1', 'phone2', 'phone3', 'phone4', 'phone5']].isin(valid_phone_set).any(axis=1)]

            removed_dupes_df = remove_phone_dupes(output_df)
            final_df = clean_contact_id_deal_id(removed_dupes_df, valid_id_set)

            # Export to save path
            export_output(final_df, list_file, save_path)
        
        print("Sucessfully processed all files")

    except Exception as e:
        print(f"An error occurred: {e}")
        raise RuntimeError   

if __name__ == "__main__":
    main()
