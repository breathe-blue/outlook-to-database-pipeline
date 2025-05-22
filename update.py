# modules required
import subprocess
from os import Path
from os.path import join
import pandas as pd
from sqlalchemy import create_engine, inspect, text
from extract import data, logging, file_dir

# sort data from all downloaded files
def data_sort(download_path):

    grouped_data = {}
    sheet_names = []
    download_path = Path(download_path)

    # list all files in the folder
    today_files = [file for file in download_path]

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
            logging.info(f"Failed reading file {file}: {e}")

    for sheet in grouped_data:
        grouped_data[sheet] = pd.concat(grouped_data[sheet], ignore_index=True)

    logging.info("Sheets identified:", sheet_names)
    return grouped_data, sheet_names


# sync data into the database table
def table_sync(sheet_name, data_frame, engine, id_field):
    insert_count = 0
    update_count = 0
    failed_rows = []

    try:
        inspector = inspect(engine)
        db_tables = [table.lower() for table in inspector.get_table_names(schema='public')]

        sheet = sheet_name.lower()
        id_field = id_field.lower()
        data_frame.columns = data_frame.columns.str.lower()

        if sheet not in db_tables:
            logging.info(f"Sheet '{sheet}' does not match any table in the database.")
            return insert_count, update_count, failed_rows

        logging.info(f"Table '{sheet}' found in the database.")

        table_columns = [col['name'].lower() for col in inspector.get_columns(sheet, schema='public')]

        if set(data_frame.columns) != set(table_columns):
            logging.info(f"Column mismatch in '{sheet}'. Expected: {table_columns}, Found: {list(data_frame.columns)}")
            return insert_count, update_count, failed_rows

        if id_field not in data_frame.columns:
            logging.warning(f"'{id_field}' field not found in DataFrame. Cannot perform upsert.")
            return insert_count, update_count, failed_rows

        # fetch existing IDs from the database
        with engine.connect() as conn:
            existing_ids_result = conn.execute(text(f"SELECT {id_field} FROM public.{sheet}"))
            existing_ids = {row[0] for row in existing_ids_result}

        # separate rows to insert or update
        df_existing = data_frame[data_frame[id_field].isin(existing_ids)]
        df_new = data_frame[~data_frame[id_field].isin(existing_ids)]

        # update existing rows
        with engine.begin() as conn:
            for _, row in df_existing.iterrows():
                try:
                    set_clause = ", ".join([f"{col} = :{col}" for col in data_frame.columns if col != id_field])
                    update_query = text(
                        f"UPDATE public.{sheet} SET {set_clause} WHERE {id_field} = :{id_field}"
                    )
                    conn.execute(update_query, row.to_dict())
                    update_count += 1
                except Exception as e:
                    logging.error(f"Failed to update row with {id_field}={row[id_field]}: {e}")
                    failed_rows.append(row.to_dict())

        # insert new rows
        if not df_new.empty:
            try:
                df_new.to_sql(sheet, engine, schema='public', if_exists='append', index=False)
                insert_count = len(df_new)
            except Exception as e:
                logging.error(f"Failed to insert new rows: {e}")
                failed_rows.extend(df_new.to_dict(orient='records'))

        logging.info(f"Inserted: {insert_count}, Updated: {update_count}, Failed: {len(failed_rows)}")

    except Exception as e:
        logging.error(f"Error during table synchronization for sheet '{sheet_name}': {e}")
        failed_rows.extend(data_frame.to_dict(orient='records'))

    return insert_count, update_count, failed_rows



# main function
def main():

    download_path = join(file_dir, "file_downloads")
    grouped_data, sheet_names = data_sort(download_path)

    # config parameters
    sql_db = data['sql_db']
    database = sql_db['database']
    user = sql_db['user']
    password = sql_db['password']
    host = sql_db['host']
    port = sql_db['port']
    id_field = data['id_field']

    # database connection
    postgresql_url = f"postgresql://{user}:{password}@{host}:{port}/{database}"
    engine = create_engine(postgresql_url)
    logging.info("Database connection established.")

    for sheet in sheet_names:
        df = grouped_data[sheet]
        insert_count, update_count, failed_rows = table_sync(sheet, df, engine, id_field)
        failed = len(failed_rows)


    engine.dispose()
    logging.info("Data sync complete. Connection closed.")

    try:
        script_path = Path(__file__).parent / "notification.py"
        subprocess.run(["python", str(script_path)], check=True)
        logging.info(f"Successfully executed {script_path}")
    except Exception as e:
        logging.error(f"Error executing update.py: {str(e)}")


if __name__ == "__main__":
    main()
