# pip install -r requirements.txt

import  pandas as pd
from    sqlalchemy import create_engine
import  sys
import  re

number_of_records = 0

def connect_to_sql_server():
    server = 'g7w11206g.inc.hpicorp.net'  # Make sure to use double backslashes
    database = 'CSI'
    
    # Create a URL for the SQLAlchemy engine using Windows Authentication
    connection_url = f"mssql+pyodbc://{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes"
    
    # Create SQLAlchemy engine
    engine = create_engine(connection_url)
    return engine

def process_groups(sub_df, result_type, yearmonth=None, week_start=None):

    group_counter_out = 0
    group_counter_in = 0
    group_start = None
    key_name = None
    split_type = None
    
    # Reset the index of sub_df
    sub_df = sub_df.sort_values(by='reference_date', ascending=True).reset_index(drop=True)
    
    # Create a new column for acc_resolved_gnrl_prev and shift the acc_resolved_gnrl values by 1
    sub_df['acc_resolved_gnrl_prev'] = sub_df['acc_resolved_gnrl'].shift(1, fill_value=0)
    
    # Initialize the index
    index = 0

    # Initialize a DataFrame to accumulate the results
    accumulated_results_df = pd.DataFrame()
    
    # Run the outer loop twice: once for 'OUT' and once for 'IN'
    for split_type in ['OUT', 'IN']:
        # Reset the index for the second run
        index = 0
        prev_index = None
        
        # Outer loop to go through each row
        while index < len(sub_df):
            row = sub_df.iloc[index]
            if prev_index == index:
                index += 1
            prev_index = index

            # Check for the start of a group based on split_type
            if (split_type == 'OUT' and row['acc_resolved_gnrl'] > 0 and row['acc_resolved_gnrl_prev'] <= 0) or \
               (split_type == 'IN' and row['acc_resolved_gnrl'] < 0 and row['acc_resolved_gnrl_prev'] >= 0):
                group_start = index
                group_counter = group_counter_out if split_type == 'OUT' else group_counter_in
                group_counter += 1
                key_name = f"{row['key']}{group_counter}"
                grp = sub_df.iloc[group_start:index + 1]

                # Inner loop to find the end of the group
                while index < len(sub_df) - 1:
                    index += 1
                    next_row = sub_df.iloc[index]
                    if (split_type == 'OUT' and next_row['acc_resolved_gnrl'] <= 0 and next_row['acc_resolved_gnrl_prev'] > 0) or \
                       (split_type == 'IN' and next_row['acc_resolved_gnrl'] >= 0 and next_row['acc_resolved_gnrl_prev'] < 0):
                        break
                    
                # Call the populate_split_results function for the identified group
                subgroup = sub_df.iloc[group_start:index + 1]
                split_results_df = populate_split_results(subgroup, key_name, split_type, result_type, yearmonth, week_start)
                
                # Accumulate the results
                accumulated_results_df = pd.concat([accumulated_results_df, split_results_df], ignore_index=True)
                
                # Update the group counter
                if split_type == 'OUT':
                    group_counter_out = group_counter
                else:
                    group_counter_in = group_counter
            else:
                # Increment the index for the outer loop and continue to the next iteration
                index += 1

    return accumulated_results_df

def process_results_table(yearmonth=None, week_start=None):

    # Connect to SQL Server and fetch data
    engine = connect_to_sql_server()
    if yearmonth:
        query = f"SELECT * FROM [OPS].[GPS_tbl_DTFC_pref_results] WHERE [res_year_month] = '{yearmonth}' ORDER BY [key], [reference_date]"
        delete_query = f"DELETE FROM [OPS].[GPS_tbl_DTFC_split_results_month] WHERE [res_year_month] = '{yearmonth}'"
        result_type = "month"
        print(f"Processing for month: {yearmonth}")
    elif week_start:
        week_end = (pd.to_datetime(week_start) + pd.Timedelta(days=6)).strftime('%Y-%m-%d')
        query = f"SELECT * FROM [OPS].[GPS_tbl_DTFC_pref_results] WHERE [res_week] = '{week_start}' ORDER BY [key], [reference_date]"
        delete_query = f"DELETE FROM [OPS].[GPS_tbl_DTFC_split_results_month] WHERE [res_week] = '{week_start}'"
        result_type = "week"
        print(f"Processing for week starting: {week_start}")
    else:
        print("Either yearmonth or week_start must be provided.")
        return
    
    df = pd.read_sql(query, engine)
    num_records = len(df)
    
    # df.to_csv('raw_df.csv', index=False)
    print(f"The number of records loaded from SQL to the DataFrame is: {num_records}")

    # Get a raw DBAPI connection from the engine
    connection = engine.raw_connection()
    
    try:
        # Create a cursor and execute the DELETE statement
        cursor = connection.cursor()
        cursor.execute(delete_query)
        connection.commit()
        print(f"Old '{yearmonth}' records deleted.")
    finally:
        # Ensure the connection is closed
        connection.close()

    # Dispose the engine
    engine.dispose()
    
    # Sort the DataFrame based on 'key'
    df.sort_values(by=['key'], inplace=True)
    df.reset_index(drop=True, inplace=True)
    
    # Initialize variables
    current_key = None
    start_row = None
    end_row = None
    
    # Initialize a DataFrame to accumulate the final results
    final_accumulated_results_df = pd.DataFrame()
    
    # Loop through the rows of the DataFrame
    for index, row in df.iterrows():
        if row['key'] != current_key:
            # Process the previous group if exists
            if start_row is not None:
                sub_df = df.iloc[start_row:end_row + 1]
                accumulated_results_df = process_groups(sub_df, result_type, yearmonth, week_start)
                # Check if accumulated_results_df is not None and not empty before concatenating
                if accumulated_results_df is not None and not accumulated_results_df.empty:
                    # Accumulate the results
                    final_accumulated_results_df = pd.concat([final_accumulated_results_df, accumulated_results_df], ignore_index=True)
            # Update the current key and reset the start row
            current_key = row['key']
            start_row = index
        
        end_row = index
    
    # Process the last group
    if start_row is not None:
        sub_df = df.iloc[start_row:end_row + 1]
        accumulated_results_df = process_groups(sub_df, result_type, yearmonth, week_start)
        # Check if accumulated_results_df is not None and not empty before concatenating
        if accumulated_results_df is not None and not accumulated_results_df.empty:
            # Accumulate the results
            final_accumulated_results_df = pd.concat([final_accumulated_results_df, accumulated_results_df], ignore_index=True)

    final_accumulated_results_df.to_csv('final_accumulated_results_df.csv', index=False)

     # Connect to SQL Server
    engine = connect_to_sql_server()
    res_len = len(final_accumulated_results_df)
    print(f"The number of records being uploaded to SQL is: {res_len}")

    # Insert the content of final_accumulated_results_df into the SQL Server table
    final_accumulated_results_df.to_sql(name='GPS_tbl_DTFC_split_results_month', con=engine, schema='OPS', if_exists='append', index=False)

    # Dispose the engine
    engine.dispose()   

def populate_split_results(arr: pd.DataFrame, group_name: str, split_type: str, result_type: str, yearmonth=None, week_start=None) -> pd.DataFrame:
    # Reset the index of arr
    arr.reset_index(drop=True, inplace=True)
    
    # Initialize the dictionary
    used_qty = {}
    
    # Initialize the results DataFrame
    split_results_df = pd.DataFrame(columns=['res_year_month', 'res_week', 'res_type', 'reference_date', 'key', 'qty_full', 'qty_split', 'diff_date', 'diff_days'])

    # Define columns based on split_type
    qty_column = 'qty_push_out' if split_type == 'OUT' else 'qty_received'
    resolving_qty_column = 'qty_received' if split_type == 'OUT' else 'first_commit_qty'

    if result_type == "month":
        res_type = 'MONTH_OUT' if split_type == 'OUT' else 'MONTH_IN'
    elif result_type == "week":
        res_type = 'WEEK_OUT' if split_type == 'OUT' else 'WEEK_IN'
    else:
        raise ValueError("Invalid result_type provided")
    
    # Loop through the array data
    for i in range(len(arr)):
        # Use .iloc for row selection and .loc for column selection
        current_row = arr.iloc[i]

        # Check if qty_column has a value
        if pd.to_numeric(current_row.loc[qty_column], errors='coerce') > 0:
            # Initialize unresolved_qty
            unresolved_qty = current_row.loc[qty_column]
            
            # Loop to resolve the qty
            j = i
            while unresolved_qty > 0 and j < len(arr):
                # Use .iloc for row selection and .loc for column selection
                resolving_row = arr.iloc[j]

                # Check resolving_qty_column
                if resolving_row.loc[resolving_qty_column] > 0:
                    # Determine available resolving_qty after accounting for used quantity
                    available_resolving_qty = resolving_row.loc[resolving_qty_column] - used_qty.get(resolving_row.loc['reference_date'], 0)
                    
                    # If there's available resolving_qty, proceed
                    if available_resolving_qty > 0:
                        # Determine qty_split based on available_resolving_qty and unresolved_qty
                        qty_split = min(unresolved_qty, available_resolving_qty)
                        
                        # Update used quantity
                        used_qty[resolving_row.loc['reference_date']] = used_qty.get(resolving_row.loc['reference_date'], 0) + qty_split
                        unresolved_qty -= qty_split
                        
                        # Append a new row to the split_results_df
                        new_row = pd.DataFrame({
                            'res_year_month': [yearmonth if yearmonth else None],
                            'res_week': [week_start if week_start else None],
                            'res_type': [res_type],
                            'reference_date': [current_row.loc['reference_date']],
                            'key': [current_row.loc['key']],
                            'qty_full': [current_row.loc[qty_column]],
                            'qty_split': [qty_split],
                            'diff_date': [resolving_row.loc['reference_date']],
                            'diff_days': [(resolving_row.loc['reference_date'] - current_row.loc['reference_date']).days]
                        })

                        split_results_df = pd.concat([split_results_df, new_row], ignore_index=True)
                j += 1

    return split_results_df

def is_valid_date_format(date_str):
    """Check if the provided date string is in the yyyy-mm-dd format and represents a Monday."""
    # Check the format using a regex
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", date_str):
        return False
    
    # Convert to datetime and check if it's a Monday (0 represents Monday in pandas)
    date_obj = pd.to_datetime(date_str, errors='coerce')
    if date_obj is pd.NaT or date_obj.weekday() != 0:
        return False
    
    return True

# Call the process_results_table function
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python out-in_split_PRD_02.py <yearmonth> OR python out-in_split_PRD_02.py --week <week_start>")
        sys.exit(1)
    
    if sys.argv[1] == "--week":
        week_starts = sys.argv[2].split(',')
        for week_start in week_starts:
            if not is_valid_date_format(week_start):
                print(f"Error: Invalid date format for {week_start}. Please provide the week start date(s) in the yyyy-mm-dd format, and it should be a Monday.")
                continue
            
            process_results_table(week_start=week_start)
            print(f"\033[92mWeekly push-OUT split results processed for {week_start}.\033[0m")
    else:
        yearmonths = sys.argv[1].split(',')
        for yearmonth in yearmonths:
            process_results_table(yearmonth=yearmonth)
            print(f"\033[92mMonthly push-OUT split results processed for {yearmonth}.\033[0m")
