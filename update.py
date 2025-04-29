import os
import sys
from os.path import join
import pandas as pd
from sqlalchemy import create_engine, inspect
import arcpy, datetime
from extract import data, logging, file_dir

def recent_file(attach_path):
    logging.info(f"Looking for CSV/XLSX/XLS files modified today in: {attach_path}")
    today = datetime.date.today()
    today_files = []

    try:
        for filename in os.listdir(attach_path):
            if filename.lower().endswith((".csv", ".xlsx", ".xls")):
                full_path = os.path.join(attach_path, filename)
                modified_time = datetime.date.fromtimestamp(os.path.getmtime(full_path))
                if modified_time == today:
                    today_files.append(full_path)
        logging.info(f"Found {len(today_files)} file(s) modified today.")
    except Exception as e:
        logging.error("Error while scanning the attachment directory.")
        logging.error(e)

    return today_files


def data_sort(today_files):
    grouped_data = {}
    sheet_names = []

    for file in today_files:
        try:
            if file.endswith(('.xlsx', '.xls')):
                excel_file = pd.ExcelFile(file)
                for sheet in excel_file.sheet_names:
                    df = excel_file.parse(sheet)
                    df.columns = df.columns.str.lower().str.replace(' ', '_')
                    sheet_lower = sheet.lower()
                    
                    if sheet_lower not in grouped_data:
                        grouped_data[sheet_lower] = [df]
                        sheet_names.append(sheet_lower)
                    else:
                        grouped_data[sheet_lower].append(df)

            elif file.endswith('.csv'):
                df = pd.read_csv(file)
                df.columns = df.columns.str.lower().str.replace(' ', '_')
                sheet_lower = 'csv'
                
                if sheet_lower not in grouped_data:
                    grouped_data[sheet_lower] = [df]
                    sheet_names.append(sheet_lower)
                else:
                    grouped_data[sheet_lower].append(df)

        except Exception as e:
            logging.error(f"Failed reading file {file}: {e}")

    for sheet in grouped_data:
        grouped_data[sheet] = pd.concat(grouped_data[sheet], ignore_index=True)

    return grouped_data, sheet_names
    


def table_sync(sheet_name, data_frame, engine):
    try:
        
        inspector = inspect(engine)

        db_tables = [table.lower() for table in inspector.get_table_names(schema='sde')]
        sheet_name_lower = sheet_name.lower()

        if sheet_name_lower not in db_tables:
            logging.error(f"Sheet '{sheet_name}' does not match any table in the database.")
            sys.exit(1)

        logging.info(f"Table '{sheet_name_lower}' found in the database.")

        table_columns = [col['name'].lower() for col in inspector.get_columns(sheet_name_lower, schema='sde')]

        data_frame.columns = data_frame.columns.str.lower()

        if set(data_frame.columns) != set(table_columns):
            logging.error(f"Columns mismatch for sheet '{sheet_name}'. Expected columns: {table_columns}, Found: {list(data_frame.columns)}")
            sys.exit(1)

        logging.info(f"Sheet '{sheet_name}' matched successfully with the database table columns.")

        data_frame.to_sql(sheet_name_lower, engine, schema='sde', if_exists='replace', index=False)

        logging.info(f"Successfully updated table '{sheet_name_lower}'.")

    except Exception as e:
        logging.error(f"Error during table synchronization for sheet '{sheet_name}': {e}")
        sys.exit(1)


def main():

    today_files = recent_file(attach_path)
    grouped_data, sheet_names = data_sort(today_files)


    sql_db = data['sql_db']
    database = sql_db['database']
    user = sql_db['user']
    password = sql_db['password']
    port = sql_db ['port']
    host = sql_db ['host']

    attach_path = join(file_dir, 'file_downloads')
    logging.info(f"Json file path: {attach_path}")

    postgresql_url = f"postgresql://{user}:{password}@{host}:{port}/{database}"
    engine = create_engine(postgresql_url)
    arcpy.AddMessage("DB Connection Created")
    table_sync(sheet_names, grouped_data, engine)

    engine.dispose()


    return None