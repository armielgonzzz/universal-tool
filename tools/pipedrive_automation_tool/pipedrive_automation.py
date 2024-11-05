import os
import pandas as pd
import pymysql
from dotenv import load_dotenv
from tqdm import tqdm

def read_file(path: str) -> pd.DataFrame:

    if path.endswith('.csv'):
        return pd.read_csv(path, low_memory=False)
    
    elif path.endswith('.xlsx'):
        return pd.read_excel(path)
    
    else:
        return None

def connect_to_db():

    host = os.getenv('DB_HOST')
    user = os.getenv('DB_USER')
    name = os.getenv('DB_NAME')
    password = os.getenv('DB_PASSWORD')

    connection = pymysql.connect(
        host=host,
        user=user,
        database=name,
        password=password,
    )

    return connection

def get_serials(
        database_id: str,
        df: pd.DataFrame,
        cursor,
        row,
        i
    ) -> None:

    serials_query = f"""
    SELECT
        c.deal_id,
        GROUP_CONCAT(DISTINCT s.serial_number SEPARATOR ' | ') AS serial_numbers
    FROM
        contact_serial_numbers s
    LEFT JOIN
        contacts c ON s.contact_id = c.id
    WHERE 1=1
        AND s.contact_id IN ({database_id})
        AND s.serial_number NOT LIKE '%MUS%'
        AND s.serial_number NOT LIKE '%CMS%'
        AND c.deal_id IS NOT NULL
    GROUP BY
        c.deal_id;
    """

    # Execute query and fetch result
    cursor.execute(serials_query)
    result = cursor.fetchone()

    # Collect serial numbers from the query
    if result and result[1]:
        unique_serials = set(result[1].split(" | "))
    else:
        unique_serials = set()

    # Process the 'Deal - Serial Number' column in the DataFrame to add unique serials
    serial_column_value = row.get('Deal - Serial Number')
    if pd.notna(serial_column_value):
        # Split serials if there's a " | " separator and add only unique ones
        for serial in str(serial_column_value).split(" | "):
            if serial not in unique_serials:
                unique_serials.add(serial)

    # Join the unique serials and store in 'new_serials'
    df.loc[i, 'new_serials'] = " | ".join(unique_serials)

def get_mailing_address(
        database_id: str,
        df: pd.DataFrame,
        cursor,
        row,
        i
    ) -> None:

    address_query = f"""
    SELECT
        address,
        city,
        state,
        postal_code
    FROM
        contact_skip_traced_addresses
    WHERE 1=1
        AND contact_id IN ({database_id})
        AND address IS NOT NULL
        AND city IS NOT NULL
        AND state IS NOT NULL
        AND postal_code IS NOT NULL
    LIMIT 1;
    """    

    address_column_value = row.get('Person - Mailing Address')
    if pd.isna(address_column_value):
    
        cursor.execute(address_query)
        result = cursor.fetchone()

        if result:
            new_address = f"{result[0]}, {result[1]}, {result[2]}, {result[3]}, USA".upper()
        else:
            new_address = None

        df.loc[i, 'Person - Mailing Address'] = new_address

def get_phone_number(
        database_id: str,
        df: pd.DataFrame,
        cursor,
        row,
        i
    ) -> None:

    phone_number_query = f"""
    SELECT 
        DISTINCT phone_number
    FROM
        contact_phone_numbers
    WHERE 1=1
        AND contact_id IN ({database_id})
        AND phone_number IS NOT NULL
        AND phone_index IS NOT NULL
    ORDER BY
        phone_index;
    """

    phone_number_column_value = row.get('Person - Phone 1')
    if pd.isna(phone_number_column_value):

        cursor.execute(phone_number_query)
        result = cursor.fetchall()
        for phone_index, phone in enumerate(result, start=1):
            df.loc[i, f'Person - Phone {phone_index}'] = phone

def get_email_address(
        database_id: str,
        df: pd.DataFrame,
        cursor,
        row,
        i
    ) -> None:

    pass

def split_id(id):
    return ", ".join(id.split(" | "))

def main(files: tuple, save_path: str):

    try:
        load_dotenv(dotenv_path='./misc/.env')

        # for i, file in enumerate(files, start=1):
        # print(f"Processing {os.path.basename(file)}")

        df = read_file('tools\pipedrive_automation_tool\deals-13519371-40929.xlsx')
        df['new_id'] = df['Deal - Unique Database ID'].apply(split_id)
        df['new_serials'] = ''
        connection = connect_to_db()

        with connection.cursor() as cursor:
            for i, row in tqdm(df.iterrows(), total=df.shape[0], unit='entry'):

                database_id = row.get('new_id')
                get_serials(database_id, df, cursor, row, i)
                get_mailing_address(database_id, df, cursor, row, i)
                get_phone_number(database_id, df, cursor, row, i)
                get_email_address(database_id, df, cursor, row, i)

        print("Successfully Processed All Files")
        df.to_csv('test_output.csv', index=False)

    except Exception as e:
        print(f"An error occurred: {e}")
        raise RuntimeError
    
    finally:
        connection.close()

if __name__ == "__main__":
    main(None, None)
