import pandas as pd
import json
import warnings
import os
import re
import yaml
 
warnings.simplefilter(action='ignore', category=UserWarning)
 
def load_mapping(filename):
    with open(filename, 'r', encoding='utf-8') as f:
        #return json.load(f)
        return yaml.safe_load(f)
 
def filter_worksheets_by_mode(worksheets, mode, mapping):
    result = []
    for ws in worksheets:
        for key, value in mapping.items():
            # Ensure correct matching against mode
            if mode == "All":
                if ws in value['aliases']:
                    result.append(ws)
                    break
            elif value['type'] == mode and ws in value['aliases']:
                result.append(ws)
                break
    return result
 
def process_excel_data(dfs, sheet_name, entity_name):
    def detect_entities_and_extract(df, entity_name):
        # Detect presence of indicative keywords
        indicative_keywords = ['Indicative adjusted', '示意性调整后', "CNY'000", "人民币千元"]
        keyword_occurrences = sum([df.to_string().count(keyword) for keyword in indicative_keywords])
 
        # Check if multiple entities are likely based on the occurrences of keywords
        is_multiple = keyword_occurrences > 1
 
        #print('detect_entities_and_extract:', 'multiple' if is_multiple else 'single', df)
 
        # Extract the entity table with the logic adjusted
        return ('multiple' if is_multiple else 'single', extract_entity_table(df, entity_name, single=not is_multiple))
 
    def extract_entity_table(df, entity_name, single=False):
        # Find the first and last non-empty column
        first_column_with_data = None
        last_column_with_data = None
 
        for i in range(len(df.columns)):
            if df.iloc[:, i].notnull().any():
                if first_column_with_data is None:
                    first_column_with_data = i
                last_column_with_data = i
            else:
                if first_column_with_data is not None:
                    break
 
        if first_column_with_data is None:
            return None
 
        # Slice the dataframe up to the last column with data
        df = df.iloc[:, :last_column_with_data + 1]
 
        result_df = None
        if single:
            result_df = df.reset_index(drop=True)
        else:
            # Try finding entity name related rows or default start
            start_idx = None
            if entity_name:
                # Prepare combinations from the entity name
                components = entity_name.split()
                combinations = {entity_name}
 
                # Create combinations
                for i in range(len(components)):
                    for j in range(i, len(components)):
                        combination = ' '.join(components[i:j+1])
                        combinations.add(combination)
 
                for i in range(len(df)):
                    if df.iloc[i].astype(str).str.contains('|'.join(combinations), na=False).any():
                        start_idx = i
                        break
           
            # If entity not found, consider a logical start point to slice the table
            if start_idx is None:
                start_idx = 0  # Fallback if entity_name not found
       
            # Collect rows until an empty row is encountered
            result_rows = []
            for i in range(start_idx, len(df)):
                if df.iloc[i].isnull().all():
                    break
                result_rows.append(df.iloc[i])
           
            result_df = pd.DataFrame(result_rows).reset_index(drop=True)
 
        # Drop empty columns, if any
        #if result_df is not None:
        #    result_df.dropna(axis=1, how='all', inplace=True)
 
        return result_df
   
    def preprocess_date(date_str):
        if isinstance(date_str, str):
            # English date formats
            if date_str.startswith('FY'):
                # Convert 'FY20' to '2020-12-31' (end of fiscal year assumed as Dec 31)
                year = int('20' + date_str[2:])
                return f'{year}-12-31'
            elif 'M' in date_str:
                # Handle '9M22' or similar formats
                try:
                    months, year_suffix = date_str.split('M')
                    year = int('20' + year_suffix)
                    end_month = int(months)
                    if end_month <= 12:
                        return pd.to_datetime(f'{year}-{end_month + 1}-01') - pd.Timedelta(days=1)
                except ValueError:
                    pass
 
            # Chinese date formats
            # Match full dates like '2021年12月31日'
            match_full_date = re.match(r'(\d{4})年(\d{1,2})月(\d{1,2})日$', date_str)
            if match_full_date:
                year, month, day = match_full_date.groups()
                return f'{year}-{month.zfill(2)}-{day.zfill(2)}'
 
            # Match dates like '2024年1-5月'
            match_period = re.match(r'(\d{4})年(\d{1,2})-(\d{1,2})月$', date_str)
            if match_period:
                year, start_month, end_month = match_period.groups()
                last_day = pd.to_datetime(f'{year}-{end_month}-01') + pd.tseries.offsets.MonthEnd(0)
                return last_day.strftime('%Y-%m-%d')
 
            # Match year only dates like '2021年'
            match_year_only = re.match(r'(\d{4})年$', date_str)
            if match_year_only:
                year = match_year_only.group(1)
                return f'{year}-12-31'
 
        return date_str
   
    def find_columns(result_table):
 
        result_table = result_table.astype(str).apply(lambda x: x.str.strip())
        #print(result_table[:5])
 
        value_column_name = None
        description_column_name = None
        value_column_number = None
 
        indicative_row_idx = None
        indicative_col_index = None
 
        for idx, row in result_table.iterrows():
            if row.apply(lambda x: bool(re.match(r'^\s*(Indicative adjusted|示意性调整后)\s*$', str(x), re.IGNORECASE))).any():
                indicative_row_idx = idx
                indicative_col_name = row[row.apply(lambda x: bool(re.match(r'^\s*(Indicative adjusted|示意性调整后)\s*$', str(x), re.IGNORECASE)))].index[0]
                indicative_col_index = result_table.columns.get_loc(indicative_col_name)
                break
        #print(indicative_col_index)
 
        if indicative_row_idx is not None and indicative_col_index is not None:
            # Starting from the indicative column index, iterate to the right until an empty or NaN is found
            for col_index in range(indicative_col_index, len(result_table.columns)):
                if pd.isna(result_table.iloc[indicative_row_idx, col_index]) or result_table.iloc[indicative_row_idx, col_index] == '':
                    last_column_with_data = col_index - 1
                    break
            else:
                last_column_with_data = len(result_table.columns) - 1  # in case there's no empty cell after the indicative col
           
            date_row_idx = indicative_row_idx + 1
            if date_row_idx < len(result_table.index):
                #print(date_row_idx, indicative_row_idx, indicative_col_index, last_column_with_data)
                #print(result_table.iloc[date_row_idx, indicative_col_index:last_column_with_data + 1])
                # Preprocess the dates column for custom formats
                preprocessed_dates = result_table.iloc[date_row_idx, indicative_col_index:last_column_with_data + 1].apply(preprocess_date)
                #preprocessed_dates = preprocessed_dates.dropna()  # Drop NaN values
               
                # Convert to datetime, coercing errors to NaT (Not a Time)
                preprocessed_dates = pd.to_datetime(preprocessed_dates, errors='coerce')
 
                # Drop NaT values (non-dates)
                preprocessed_dates = preprocessed_dates.dropna()
 
                # Print the resulting dates
                #print('preprocessed_dates:', preprocessed_dates)
 
                try:
                    date_columns = preprocessed_dates.apply(pd.to_datetime, errors='raise').dropna()  # Raise error if invalid
                    #print(result_table)
                   
                    if not date_columns.empty:
                        most_recent_date_idx = date_columns.idxmax()
                       
                        if pd.notna(most_recent_date_idx):
                            value_column_name = date_columns[most_recent_date_idx].strftime('%Y-%m-%d')
                            value_column_index = result_table.columns.get_loc(most_recent_date_idx)
                            value_column_number = value_column_index + 1
 
                except Exception as e:
                    # raise ValueError(f"There was an error with the date conversion: {str(e)}")
                    # Extract column name where the error occurred
                    error_column_name = preprocessed_dates.name if hasattr(preprocessed_dates, 'name') else 'unknown'
                    result_table = result_table.drop(columns=[error_column_name])  # Drop the problematic column
                    #print(f"Dropped column due to error in date conversion: {error_column_name}. Error: {str(e)}")
               
        description_col_original = None
        for col in result_table.columns:
            if result_table[col].str.contains(r"CNY'000|人民币千元", case=False, na=False).any():
                description_col_original = col
                description_column_name = result_table.iloc[0, result_table.columns.get_loc(col)]
                break
 
        if value_column_name and description_col_original is not None:
            result_df = result_table.iloc[date_row_idx + 1:, [result_table.columns.get_loc(description_col_original), value_column_index]]
            result_df.columns = [description_column_name, value_column_name]
            result_df[value_column_name] = result_df[value_column_name].astype(float)
            result_df = result_df[result_df[value_column_name] != 0]
 
            # Check all values in a specific column or set of columns
            contains_cny = False
            contains_chinese_cny = False
 
            # Iterate over all elements in the row to find the key phrase
            for element in result_table.iloc[date_row_idx]:
                # Convert each element to a string and check
                if isinstance(element, str):
                    if "CNY'000" in element:
                        contains_cny = True
                    if "人民币千元" in element:
                        contains_chinese_cny = True
 
            if contains_cny or contains_chinese_cny:
                #print('Triggered')
                result_df[value_column_name] *= 1000
 
            #print("After multiplying:")
            #print(result_df[value_column_name])
 
            return result_df.reset_index(drop=True), value_column_index
        else:
            return None, None
 
    # Fetch the dataframe from the dictionary
    df = dfs.get(sheet_name)
    if df is None:
        raise ValueError(f"No data found for sheet: {sheet_name}")
 
    result_type, result_table = detect_entities_and_extract(df, entity_name)
    #print('result_table:', result_table)
    if result_table is not None:
        extracted_df, value_col_num = find_columns(result_table)
        return result_type, extracted_df, value_col_num
    else:
        return result_type, None, None
   
def determine_result_type(sheet_data):
    # Check the number of occurrences of the keywords in the sheet
    indicative_keywords = ['Indicative adjusted', '示意性调整后', "CNY'000", "人民币千元"]
    occurrences = sum([sheet_data.to_string().count(keyword) for keyword in indicative_keywords])
   
    # Determine if the sheet is single or multiple based on occurrences
    return 'multiple' if occurrences > 1 else 'single'
 
def extract_data_from_excel(databook_path, entity_name, mode="All"):
    """
    Extract data from Excel file and determine language.
    
    Args:
        databook_path: Path to Excel file
        entity_name: Name of entity to extract
        mode: Filter mode ('All', 'Assets', 'Liabilities', 'Equity', 'Income', 'Expenses')
    
    Returns:
        Tuple of (final_dfs, final_workbook_list, overall_result_type, report_language)
    """
    mapping_file = os.path.join(os.path.dirname(__file__), 'mappings.yml')
    xls = pd.ExcelFile(databook_path)
    all_sheets = xls.sheet_names
    mapping = load_mapping(mapping_file)
    filtered_sheets = filter_worksheets_by_mode(all_sheets, mode, mapping)

    raw_dfs = {sheet: pd.read_excel(databook_path, sheet_name=sheet, engine='openpyxl') for sheet in filtered_sheets}

    final_dfs = {}
    final_workbook_list = []
    single_count = 0
    multiple_count = 0

    english_count = 0
    chinese_count = 0

    for sheet in filtered_sheets:
        result_type, extracted_df, value_col_num = process_excel_data(raw_dfs, sheet, entity_name)
       
        if extracted_df is not None and not extracted_df.dropna().empty and value_col_num is not None:
            final_dfs[sheet] = extracted_df
            final_workbook_list.append(sheet)

            # Determine result type using the new logic
            sheet_result_type = determine_result_type(raw_dfs[sheet])
            if sheet_result_type == 'multiple':
                multiple_count += 1
            else:
                single_count += 1

            # Check report language
            if 'Indicative adjusted' in raw_dfs[sheet].to_string():
                english_count += 1
            elif '示意性调整后' in raw_dfs[sheet].to_string():
                chinese_count += 1

    # Avoid division by zero and determine the overall result type
    if len(final_workbook_list) == 0:
        overall_result_type = 'None'
    else:
        overall_result_type = 'multiple' if multiple_count > single_count else 'single'

    # Determine report language with better logic
    if len(final_workbook_list) == 0:
        report_language = None
    elif english_count / len(final_workbook_list) >= 0.8:
        report_language = 'Eng'
    elif chinese_count / len(final_workbook_list) >= 0.8:
        report_language = 'Chi'
    else:
        # Default to English if unclear
        report_language = 'Eng'

    return final_dfs, final_workbook_list, overall_result_type, report_language