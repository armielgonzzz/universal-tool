import os
import pandas as pd
from datetime import datetime, timedelta

def read_file(path: str):

    if path.endswith('.csv'):
        return pd.read_csv(path, low_memory=False)

    elif path.endswith('.xlsx'):
        return pd.read_excel(path)
    
    else:
        raise ValueError("Invalid file format: Please provide a .csv or .xlsx file.")

    
def apply_mask(df: pd.DataFrame, mask, output_value) -> None:
    df.loc[mask, 'reason_for_removal'] = df.loc[mask, 'reason_for_removal'].apply(
        lambda lst: lst + [output_value] if isinstance(lst, list) else [output_value])


def check_date(df: pd.DataFrame, col: str) -> pd.DataFrame:
    current_date = datetime.now().date()
    seven_days_ago = current_date - timedelta(days=7)
    df[col] = pd.to_datetime(df[col], errors='coerce').dt.date
    mask = (df[col] >= seven_days_ago) & (df[col] <= current_date)

    return mask


def apply_all_filters(df: pd.DataFrame) -> pd.DataFrame:
    
    df['reason_for_removal'] = [[] for _ in range(len(df))]
    
    if 'Rolling 30 Days Rvm Count' in df.columns:
        df['Rolling 30 Days Rvm Count'] = pd.to_numeric(df['Rolling 30 Days Rvm Count'], errors='coerce')
    
    if 'Rolling 30 Days Text Marketing Count' in df.columns:
        df['Rolling 30 Days Text Marketing Count'] = pd.to_numeric(df['Rolling 30 Days Text Marketing Count'], errors='coerce')
    
    if 'Rolling 30 Days Max Outbound Count' in df.columns:
        df['Rolling 30 Days Max Outbound Count'] = pd.to_numeric(df['Rolling 30 Days Max Outbound Count'], errors='coerce')

    if 'in_pipedrive' in df.columns:
        in_pipedrive_mask = df['in_pipedrive'].str.upper() == 'Y'
        apply_mask(df, in_pipedrive_mask, 'in_pipedrive is Y')

    if 'rc_pd' in df.columns:
        rc_pd_mask = df['rc_pd'].str.upper() == 'YES'
        apply_mask(df, rc_pd_mask, 'rc_pd is Yes')

    if 'type' and 'carrier_type' in df.columns:
        type_ctype_mask = (df['type'].str.upper() == 'LANDLINE') & (df['carrier_type'].str.upper() == 'LANDLINE')
        apply_mask(df, type_ctype_mask, 'Both type & carrier_type are Landline')

    if 'text_opt_in' in df.columns:
        text_opt_in_mask = df['text_opt_in'].str.upper() == 'NO'
        apply_mask(df, text_opt_in_mask, 'text_opt_in is No')

    if 'contact_deal_id' in df.columns:
        contact_deal_id_mask = df['contact_deal_id'].notna()
        apply_mask(df, contact_deal_id_mask, 'contact_deal_id Not Empty')

    if 'contact_deal_status' in df.columns:
        contact_deal_status_mask = df['contact_deal_status'].notna()
        apply_mask(df, contact_deal_status_mask, 'contact_deal_status Not Empty')

    if 'contact_person_id' in df.columns:
        contact_person_id_mask = df['contact_person_id'].notna()
        apply_mask(df, contact_person_id_mask, 'contact_person_id Not Empty')

    if 'phone_number_deal_id' in df.columns:
        phone_number_deal_id_mask = df['phone_number_deal_id'].notna()
        apply_mask(df, phone_number_deal_id_mask, 'phone_number_deal_id Not Empty')

    if 'phone_number_deal_status' in df.columns:
        phone_number_deal_status_mask = df['phone_number_deal_status'].notna()
        apply_mask(df, phone_number_deal_status_mask, 'phone_number_deal_status Not Empty')

    if 'RVM - Last RVM Date' in df.columns:
        last_rvm_mask = check_date(df, 'RVM - Last RVM Date')
        apply_mask(df, last_rvm_mask, 'RVM - Last RVM Date - last 7 days from tool run date')

    if 'Latest Text Marketing Date (Sent)' in df.columns:
        last_marketing_text_mask = check_date(df, 'Latest Text Marketing Date (Sent)')
        apply_mask(df, last_marketing_text_mask, 'Latest Text Marketing Date (Sent) - last 7 days from tool run date')

    if 'Rolling 30 Days Max Outbound Count' and 'Rolling 30 Days Text Marketing Count' in df.columns:
        rolling_days_mask = (df['Rolling 30 Days Max Outbound Count'] + df['Rolling 30 Days Text Marketing Count']) >= 3
        apply_mask(df, rolling_days_mask, 'Rolling 30 Days Max Outbound Count and Rolling 30 Days Text Marketing Count - total >= 3')

    if 'Deal - ID' in df.columns:
        deal_id_mask = df['Deal - ID'].notna()
        apply_mask(df, deal_id_mask, 'Deal - ID Not Empty')

    if 'Deal - Text Opt-in' in df.columns:
        deal_text_opt_in = df['Deal - Text Opt-in'].str.upper().str.contains('NO', na=False)
        apply_mask(df, deal_text_opt_in, 'Deal - Text Opt-in is No')
    
    # Deduplication
    df['reason_length'] = df['reason_for_removal'].apply(len)
    df_longest_reason = df.loc[df.groupby('phone_number', sort=False)['reason_length'].idxmax()]
    df_longest_reason = df_longest_reason.drop(columns=['reason_length'])

    return df_longest_reason


def export_output(df: pd.DataFrame, file_path: str, save_path: str) -> None:

    filename = os.path.basename(file_path)

    df['reason_for_removal'] = df['reason_for_removal'].apply(
        lambda lst: ', '.join(lst) if isinstance(lst, list) else lst
    )

    output_df = df[[
        'phone_number',
        'contact_id',
        'carrier_type',
        'full_name',
        'first_name',
        'last_name',
        'target_county',
        'target_state',
        'phone_index',
        'time_zone',
        'reason_for_removal'
    ]]

    if filename.endswith('.csv'):
        output_df.to_csv(f"{save_path}/(With Cleanup Tagging) {filename}", index=False)
    
    elif filename.endswith('.xlsx'):
        output_df.to_excel(f"{save_path}/(With Cleanup Tagging) {filename}", index=False)
    
    else:
        print("No output generated. Invalid file format")


def main(files: tuple, save_path: str):

    try:
        # input_path = 'data'
        # file_list = os.listdir(input_path)

        for file in files:

            print(f"Processing file {file}")

            df = read_file(file)
            filtered_df = apply_all_filters(df)
            export_output(filtered_df, file, save_path)
        
        print("Successfully processed all files")

    except Exception as e:
        print(f"An error occurred: {e}")
        raise RuntimeError

if __name__ == "__main__":
    main()
