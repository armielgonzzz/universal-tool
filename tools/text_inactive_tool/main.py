import os
import pandas as pd
from sqlalchemy import create_engine
from dotenv import load_dotenv
from sql_queries import *
from get_pipedrive_data import main as update_pipedrive
from follow_up import process_fu

def read_cm_live_db() -> 'tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame | None]':

    try:
        # Create database engine
        host = os.getenv('DB_HOST')
        user = os.getenv('DB_USER')
        name = os.getenv('DB_NAME')
        password = os.getenv('DB_PASSWORD')
        engine = create_engine(f'mysql+mysqlconnector://{user}:{password}@{host}/{name}')

        print(f'Reading Community Minerals Database')

        # Execute query and fetch the data into a Pandas Dataframe
        phone_number_df = pd.read_sql_query(phone_number_query, engine)
        emaiL_address_df = pd.read_sql_query(email_address_query, engine)
        serial_numbers_df = pd.read_sql_query(serial_numbers_query_mysql, engine)
        cm_db_df = pd.read_sql_query(cm_db_query, engine)

        # Change data type of phone number to int
        phone_number_df['phone_number'] = phone_number_df[phone_number_df['phone_number']\
                                                          .str.contains(r'^[0-9]+$', na=False)]\
                                                            ['phone_number'].astype('Int64')

        return phone_number_df, emaiL_address_df, serial_numbers_df, cm_db_df

    except Exception:
        return None,None,None,None

    finally:
        engine.dispose()

def read_file(path: str) -> pd.DataFrame:

    if path.endswith('.csv'):
        return pd.read_csv(path, low_memory=False)
    
    elif path.endswith('.xlsx'):
        return pd.read_excel(path)
    
    else:
        return None

def main(files: tuple, save_path: str = './output'):

    try:
        load_dotenv(dotenv_path='misc/.env')
        # update_pipedrive()
        phone_number_df, emaiL_address_df, serial_numbers_df, cm_db_df = read_cm_live_db()
        df = read_file('./tools/text_inactive_tool/Inactive Deal Tagging 102924.xlsx')
        pipedrive_df = read_file('./data/pipedrive/pipedrive_data.csv')
        no_deal_id_final = process_fu(df, pipedrive_df, phone_number_df, cm_db_df)

    except Exception as e:
        print(f"An error occured: {e}")
        raise RuntimeError


if __name__ == "__main__":
    main(None, None)
