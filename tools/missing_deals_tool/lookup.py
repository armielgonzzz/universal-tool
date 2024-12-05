import os
import pandas as pd
from datetime import datetime
from .get_pipedrive_data import main as update_pipedrive

def read_file(path: str) -> pd.DataFrame:

    if path.endswith('.csv'):
        return pd.read_csv(path, low_memory=False)
    
    elif path.endswith('.xlsx'):
        return pd.read_excel(path)
    
    else:
        return None
    
def format_phone(lookup_df: pd.DataFrame) -> pd.DataFrame:
    if 'From' in lookup_df.columns:
        phone_column = 'From'
        lookup_df[phone_column] = lookup_df[phone_column].astype(str).apply(lambda x: x[1:])
        lookup_df[phone_column] = pd.to_numeric(lookup_df[phone_column], errors='coerce').astype('Int64').astype(str)
    else:
        phone_column = 'ANI'
        lookup_df[phone_column] = lookup_df[phone_column].astype(str).apply(lambda x: x[1:] if len(x) == 11 else x)
        lookup_df[phone_column] = pd.to_numeric(lookup_df[phone_column], errors='coerce').astype('Int64').astype(str)
    
    return lookup_df, phone_column

def format_pipedrive_data(pipedrive_df: pd.DataFrame) -> dict:

    print("Formatting Pipedrive data")

    # Ensure phone_number column is a string and handle missing values
    pipedrive_df['phone_number'] = pipedrive_df['phone_number'].fillna('').astype(str)

    # Remove rows with empty or whitespace-only phone numbers
    pipedrive_df = pipedrive_df[pipedrive_df['phone_number'].str.strip() != '']

    # Split phone numbers on commas, remove duplicates, and preserve the original order
    pipedrive_df['phone_number'] = pipedrive_df['phone_number'].str.split(',')
    pipedrive_df['phone_number'] = pipedrive_df['phone_number'].apply(
        lambda x: sorted(set(x), key=x.index)
    )

    # Explode the phone_number column into individual rows
    pipedrive_final_data = pipedrive_df.explode('phone_number').reset_index(drop=True)

    # Remove non-numeric characters from phone numbers
    pipedrive_final_data['phone_number'] = pipedrive_final_data['phone_number'].str.replace(r'\D', '', regex=True)

    # Remove rows with empty phone numbers after cleaning
    pipedrive_final_data = pipedrive_final_data[pipedrive_final_data['phone_number'] != '']

    # Group by phone_number and aggregate 'Deal - ID' efficiently
    grouped_df = (
        pipedrive_final_data.groupby('phone_number')['Deal - ID']
        .agg(lambda row: " | ".join(row.astype(str).unique()))
        .reset_index()
    )
    phone_to_deal = grouped_df.set_index('phone_number')['Deal - ID'].to_dict()
    return phone_to_deal

def lookup_text(lookup_df: pd.DataFrame, phone_column: str, phone_to_deal: dict, save_path: str, i: int) -> None:
    lookup_df[['Deal ID', 'Resolved By']] = lookup_df.apply(
        lambda row: (
            phone_to_deal.get(row[phone_column], row['Deal ID']) if pd.isna(row['Deal ID']) else row['Deal ID'], 
            'Joyce Marie Gempesaw' if pd.isna(row['Deal ID']) and row[phone_column] in phone_to_deal else row['Resolved By']
        ),
        axis=1,
        result_type='expand'  # Ensure multiple values can be assigned to multiple columns
    )

    lookup_df[[phone_column, 'Deal ID', 'Resolved By']].to_excel(f'{save_path}/{i}. Text Lookup.xlsx', index=False)

def lookup_call(lookup_df: pd.DataFrame, phone_column: str, phone_to_deal: dict, save_path: str, i: int) -> None:
    lookup_df[['Deal ID', 'Resolved by', 'Resolve Date']] = lookup_df.apply(
        lambda row: (
            phone_to_deal.get(row[phone_column], row['Deal ID']) if pd.isna(row['Deal ID']) else row['Deal ID'],  # Deal ID logic
            'Joyce Marie Gempesaw' if pd.isna(row['Deal ID']) and row[phone_column] in phone_to_deal else row['Resolved by'],   # Resolved By logic
            datetime.today().strftime('%m/%d/%Y') if pd.isna(row['Deal ID']) and row[phone_column] in phone_to_deal else row['Resolve Date']  # Resolve Date logic
        ),
        axis=1,
        result_type='expand'  # Ensure multiple values can be assigned to multiple columns
    )

    lookup_df[[phone_column, 'Resolved by', 'Resolve Date', 'Deal ID']].to_excel(f'{save_path}/{i}. Call Lookup.xlsx', index=False)

def main(files: tuple, save_path: str):

    try:
        update_pipedrive()
        pipedrive_df = read_file('./data/pipedrive/pipedrive_data.csv')
        phone_to_deal_dict = format_pipedrive_data(pipedrive_df)

        for i, file in enumerate(files, start=1):

            print(f"Processing {os.path.basename(file)}")

            df = read_file(file)
            formatted_df, phone_column = format_phone(df)
            if 'Resolve Date' in formatted_df.columns:
                lookup_call(formatted_df, phone_column, phone_to_deal_dict, save_path, i) 
            else:
                lookup_text(formatted_df, phone_column, phone_to_deal_dict, save_path, i)

        print("Successfully Processed All Files")

    except Exception as e:
        print(f"An error occured: {e}")
        raise RuntimeError


if __name__ == "__main__":
    main()
