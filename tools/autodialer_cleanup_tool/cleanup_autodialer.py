import os
import re
import pandas as pd

def export_output(df: pd.DataFrame, file_path: str, save_path: str) -> None:

    filename = os.path.basename(file_path)
    if filename.endswith('.csv'):
        df.to_csv(f"{save_path}/(Clean file) {filename}", index=False)
    
    elif filename.endswith('.xlsx'):
        df.to_excel(f"{save_path}/(Clean file) {filename}", index=False)
    
    else:
        print("No output generated. Invalid file format")

def read_file(path: str):

    if path.endswith('.csv'):
        return pd.read_csv(path, low_memory=False)

    elif path.endswith('.xlsx'):
        return pd.read_excel(path)
    
    else:
        raise ValueError("Invalid file format: Please provide a .csv or .xlsx file.")
    
def is_valid_phone(phone):
    return bool(re.fullmatch(r'\d{10,15}', phone))

def main(cleanup_files: tuple, list_files: tuple, save_path: str):

    try:
        for to_clean_file in cleanup_files:

            print("Preparing List Cleaner")

            list_cleaner_df = pd.read_excel(to_clean_file,
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

            # Export to save path
            export_output(output_df, list_file, save_path)
        
        print("Sucessfully processed all files")

    except Exception as e:
        print(f"An error occurred: {e}")
        raise RuntimeError   

if __name__ == "__main__":
    main()
