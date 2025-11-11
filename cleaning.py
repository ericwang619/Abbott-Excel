import pandas as pd
from decimal import Decimal
from openpyxl.utils import get_column_letter
import time
import re

from config_headers import *


def clean_data(unit_df, form_df, sheet = data_sheet_name):
    file_name = sheet.split('/')[-1]

    print(f"Cleaning {file_name}")
    start_time = time.time()

    # data tab
    data_df = pd.read_excel(sheet, sheet_name=data_s, keep_default_na=False, dtype={project_h: str, run_h: str})

    # add new columns headers to dataframe
    add_columns(data_df, [temp_h, humidity_h, interval_h])

    # change production date to MM/DD/YYYY
    print("--Converting production and completion date formats")
    data_df[prod_date_h] = data_df[prod_date_h].astype(object)
    data_df[prod_date_h] = data_df[prod_date_h].apply(convert_dates)

    # change completion date to MM/DD/YYYY
    data_df[comp_date_h] = data_df[comp_date_h].astype(object)
    data_df[comp_date_h] = data_df[comp_date_h].apply(convert_dates)

    # extract temperature and humidity values
    print("--Converting temperature (C), humidity, and interval (D)")
    data_df[[temp_h, humidity_h]] = data_df.apply(convert_temp_humidity, axis=1)

    # populate interval (D) column
    data_df[interval_h] = data_df[dur_h].apply(convert_duration)

    # copy corresponding formula code and sources if exists
    print("--Adding formula codes and sources. Remove rows with invalid formulas")
    data_df = get_formula(data_df, form_df)

    # only keep rows with valid formula codes
    data_df = drop_invalid_formulas(data_df)

    # copy corresponding test, unit, and conversion factor if exists
    print("--Adding and applying test/unit conversions to get final results")
    first_conv_df = pd.read_excel(first_conv_file, sheet_name=first_conv_s, keep_default_na=False, skiprows=2)
    data_df = add_unit_conversions(data_df, unit_df, first_conv_df)

    # change text to numeric value if applicable
    data_df[text_h] = data_df[text_h].apply(convert_text)

    # add vitE_Factor and results column
    add_columns(data_df, [vitE_h, results_h])

    # this is the sheet/tab that holds the vitE values
    vitE_df = pd.read_excel(vitE_file, sheet_name=vitE_s, keep_default_na=False)
    vitE_map = vitE_df.set_index(form_h)[vitE_h].to_dict()

    # transform text based on conversion factor
    converted = [convert_result(row, vitE_map) for _, row in data_df.iterrows()]
    # --- Unpack into two columns safely ---
    results, vitE_vals = zip(*converted)
    data_df[results_h] = results
    data_df[vitE_h] = vitE_vals

    # drop duplicate rows
    data_df = (data_df.drop_duplicates(subset=[batch_h, project_h, prod_date_h,
                                               temp_h, ab_stage_h, interval_h,
                                               test_h, newUnit_h, results_h]).reset_index(drop=True))

    # upload changes back to the excel sheet
    print("--Adding updates to the spreadsheet")
    with pd.ExcelWriter(sheet, mode="a", if_sheet_exists="replace") as writer:

        data_df.to_excel(writer, sheet_name=updated_s, index=False)
        fit_columns(data_df, writer, updated_s)

    # create new tab with test+unit as columns and re-organize result data
    print("--Creating re-organized tab")
    new_df = consolidate(data_df)

    # upload new tab to the Excel file
    print("--Adding re-organized tab to the spreadsheet")
    # write revised data to sheet
    with pd.ExcelWriter(sheet, mode="a", if_sheet_exists="replace") as writer:

        new_df.to_excel(writer, sheet_name=consolidated_s, index=False)
        fit_columns(new_df, writer, consolidated_s)

    # print total time taken
    elapsed = (time.time() - start_time) / 60
    print(f"Finished Cleaning {file_name} in {elapsed:.3f} minutes\n")


# adjust column width to longest value in column
def fit_columns(new_df, writer, sheet_name):
    # Access the openpyxl workbook and worksheet
    worksheet = writer.sheets[sheet_name]
    # Loop through all columns and set width
    for i, col in enumerate(new_df.columns, 1):  # 1-based index
        # Compute max width between header and data
        column_len = max(
            new_df[col].astype(str).map(len).max(),  # longest cell
            len(col)  # header
        ) + 2  # padding
        worksheet.column_dimensions[get_column_letter(i)].width = column_len


# adds new column headers to dataframe
def add_columns(df, headers):
    for h in headers:
        df[h] = pd.Series(dtype=object)


# convert production and completion date formats
def convert_dates(date):

    # check if date exists
    if pd.isna(date):
        parsed_date = pd.NaT
    # check if date value is numerical
    elif isinstance(date, (int, float)):
        # Excel float
        parsed_date = pd.to_datetime(date, unit='D', origin='1899-12-30')
    else:
        # String or datetime-like
        try:
            parsed_date = pd.to_datetime(date)
        except Exception:
            parsed_date = pd.NaT

    # format as mm/dd/yyyy string
    date = parsed_date.strftime("%m/%d/%Y") if not pd.isna(parsed_date) else None

    return str(date)


# extract temperature and humidity (if possible) from storage value
def convert_temp_humidity(row):
    storage = str(row[storage_h])
    humidity = ''
    if storage == 'ROOM':
        return pd.Series([22, humidity])
    elif storage == 'REFRIG':
        return pd.Series([4, humidity])
    elif storage == "FROZEN":
        return pd.Series([-20, humidity])
    elif 'C' in storage:
        # strip 'C' from number
        i = storage.index('C')
        temp = int(storage[:i])
        humidity = storage[i+2:]
        if humidity != '':
            humidity = int(humidity)
        return pd.Series([temp, humidity])
    elif 'F' in storage:
        # calculate C from F value
        i = storage.index('F')
        temp = round((int(storage[:i])-32)/9*5, 2)
        humidity = storage[i+2:]
        if humidity != '':
            humidity = int(humidity)
        return pd.Series([temp, humidity])
    return pd.Series([storage, humidity])


# convert duration values to # days
def convert_duration(dur):
    if 'D' in dur:
        return int(dur[:dur.index('D')])
    elif 'M' in dur:
        return 30*int(dur[:dur.index('M')])
    return dur


# pull matching formula code from formula tab
def get_formula(data_df, form_df):

    # extract necessary columns from formula sheet
    form_lookup = form_df[[project_h, run_h, batch_h, conversion_formula_h, conversion_sources_h]].copy()

    # Make sure column header types match
    for col in [project_h, run_h, batch_h]:
        data_df[col] = data_df[col].astype(object)
        form_lookup[col] = form_lookup[col].astype(object)

    # left merge formula tab with data if same project, run, batch
    data_df = data_df.merge(
        form_lookup,
        how='left',
        on=[project_h, run_h, batch_h]
    )

    # After merge, rename the new columns if needed
    if conversion_formula_h in data_df.columns:
        data_df.rename(columns={conversion_formula_h: form_h}, inplace=True)
    if conversion_sources_h in data_df.columns:
        data_df.rename(columns={conversion_sources_h: sources_h}, inplace=True)

    # Fill missing cells with empty values
    data_df[[form_h, sources_h]] = data_df[[form_h, sources_h]].fillna('')

    return data_df


# remove rows with invalid/empty formulas from the dataframe
def drop_invalid_formulas(data_df):
    # Create a boolean mask for rows to drop
    pattern = '|'.join(re.escape(f) for f in invalid_formulas if f != '')
    mask_invalid = (data_df[form_h] == '') | data_df[form_h].str.contains(pattern, na=False)
    # Keep only the rows that are not invalid
    data_df = data_df[~mask_invalid].reset_index(drop=True)
    return data_df


# convert text values to number where applicable
def convert_text(text):
    try:
        text = float(text)
    finally:
        return text


# pull matching columns from conversion tab
def add_unit_conversions(data_df, unit_df, first_conv_df):
    """
    Populate test_h, newUnit_h, conv_h columns in data_df based on lookups in
    first_conv_df (primary) and unit_df (fallback). Ensures no duplicate matches
    cause incorrect assignments.
    """
    # Step 1: Prepare lookup DataFrames with unique temporary column names
    first_conv_lookup = (
        first_conv_df.rename(columns={
            first_conv_test_h: 'fc_test',
            first_conv_units_h: 'fc_units',
            first_conv_h: 'fc_conv'
        })[[analysis_h, name_h, unit_h, 'fc_test', 'fc_units', 'fc_conv']]
        .drop_duplicates(subset=[analysis_h, name_h, unit_h])
    )

    unit_lookup = (
        unit_df.rename(columns={
            conversion_test_h: 'unit_test',
            conversion_units_h: 'unit_units',
            conversion_conv_h: 'unit_conv'
        })[[analysis_h, name_h, unit_h, 'unit_test', 'unit_units', 'unit_conv']]
        .drop_duplicates(subset=[analysis_h, name_h, unit_h])
    )

    # Step 2: Merge data_df with first_conv_lookup
    merged = data_df.merge(
        first_conv_lookup,
        how='left',
        on=[analysis_h, name_h, unit_h]
    )

    # Step 3: Merge with unit_lookup as fallback (adds new columns for all rows)
    merged = merged.merge(
        unit_lookup,
        how='left',
        on=[analysis_h, name_h, unit_h]
    )

    # Step 4: Populate final columns: use first_conv values if present, else fallback
    merged[test_h] = merged['fc_test'].combine_first(merged['unit_test']).fillna('')
    merged[newUnit_h] = merged['fc_units'].combine_first(merged['unit_units']).fillna('')
    merged[conv_h] = merged['fc_conv'].combine_first(merged['unit_conv']).fillna('')

    # Step 5: Drop temporary columns
    merged.drop(columns=['fc_test', 'fc_units', 'fc_conv',
                         'unit_test', 'unit_units', 'unit_conv'],
                inplace=True, errors='ignore')

    return merged



# convert text to results using conversion factor + vitE Factor where applicable
def convert_result(row, vitE_map):

    # get text value, if not numerical, result = ?
    text = row.get(text_h)
    if not isinstance(text, (int, float)):
        return '', ''

    # get conversion factor and formula for this row
    result = Decimal(str(text))
    factor = 1
    conversion_factor = row.get(conv_h, '')
    formula = row.get(form_h, '')
    vitE_val = ''

    if conversion_factor == '':
        return '', ''

    # check for valid conversion factor
    if conversion_factor != 'Copy value':

        # perform division as needed
        terms = str(conversion_factor).split('/')

        # First term
        f0 = terms[0]

        # check if vitE factor needed
        if 'Vit E Factor' in f0:
            vitE_val = vitE_map.get(formula, '')
            if vitE_val in ['', 'Pending', None] or '?' in str(vitE_val):
                return '', vitE_val
            try:
                factor *= Decimal(str(vitE_val))
            except:
                print('vitE', vitE_val)
                raise Exception
        else:

            # if not needed, start division arithmetic if possible
            if f0[0] == '=':
                f0 = f0[1:]
            try:
                factor *= Decimal(str(float(f0)))
            except:
                return '', vitE_val

        # perform division with remaining terms
        for t in terms[1:]:
            if t == 'density':
                continue

            # check if vitE factor is needed
            elif 'Vit E Factor' in t:
                vitE_val = vitE_map.get(formula, '')
                if vitE_val in ['', 'Pending', None] or '?' in str(vitE_val):
                    return '', vitE_val
                factor /= Decimal(str(vitE_val))
            else:
                try:
                    factor /= Decimal(str(float(t)))
                except:
                    return '', vitE_val

    return float(result * factor), vitE_val


# re-organized sheet with test + units as column headers
# def consolidate(df):
#
#     # keep these from updated data
#     headers_to_keep = [form_h, project_h, run_h, batch_h, description_h, batch_type_h, batch_sub_h,
#                        manu_loc_h, prod_date_h, ab_container_h, ab_stage_h, temp_h, humidity_h, interval_h]
#
#     # copy relevant columns to new dataframe
#     new_df = df[headers_to_keep].copy().drop_duplicates()
#
#     # extract test, unit headers
#     nn_df = pd.read_excel(nutrient_file, sheet_name=nutrient_s, skiprows=1, usecols=[0,1], keep_default_na=False)
#     nn_list = set()
#     for i, row in nn_df.iterrows():
#         nn_list.add(str(row[test_h]) + ', ' + str(row[newUnit_h]))
#
#     # create new column for each header
#     for n in nn_list:
#         new_df[n] = pd.Series(dtype=object)
#
#     # remove duplicate values
#     new_df = new_df.drop_duplicates(subset=[batch_h, project_h, prod_date_h,
#                                             temp_h, ab_stage_h, interval_h])
#
#     # store updates for batch upload after looping
#     updates = []
#     total_rows = len(df)
#
#     # iterate through all updated data, copying over result values
#     for i, row in df.iterrows():
#
#         # progress tracking (since this can take a while)
#         if i != 0 and i % 15000 == 0:
#             print(f'----processing row {i} / {total_rows}')
#
#         # extract values from updated data
#         cols = [batch_h, project_h, prod_date_h, temp_h, ab_stage_h, interval_h, test_h, newUnit_h, results_h]
#         batch, project, production_date, temp, ab_stage, interval, test, units, results = row[cols]
#
#         # only copy numerical values
#         if results == '' or not isinstance(results, float):
#             continue
#
#         # find matching row in new dataframe
#         match = new_df.loc[(new_df[batch_h] == batch) &
#                            (new_df[project_h] == project) &
#                            (new_df[prod_date_h] == production_date) &
#                            (new_df[temp_h] == temp) &
#                            (new_df[ab_stage_h] == ab_stage) &
#                            (new_df[interval_h] == interval)]
#
#         # copy over results value to corresponding column
#         if len(match) > 0:
#             index = match.index[0]
#             col_name = str(test) + (', ' + str(units) if units else '')
#             updates.append((index, col_name, float(results)))
#
#     # apply all updates at once
#     for index, col_name, value in updates:
#         new_df.at[index, col_name] = value
#
#     new_df.insert(1, data_type_h, 'LIMS Test')
#
#     return new_df

def consolidate(df):
    # keep relevant headers
    headers_to_keep = [form_h, project_h, run_h, batch_h, description_h, batch_type_h, batch_category_h,
                       manu_loc_h, prod_date_h, ab_container_h, ab_stage_h, temp_h, humidity_h, interval_h]
    new_df = df[headers_to_keep].copy().drop_duplicates()

    # read nutrients and create column names
    nn_df = pd.read_excel(nutrient_file, sheet_name=nutrient_s, skiprows=1, usecols=[0,1], keep_default_na=False)
    nn_list = [str(row[test_h]) + ', ' + str(row[newUnit_h]) for _, row in nn_df.iterrows()]
    for n in nn_list:
        new_df[n] = pd.NA  # use pandas nullable type

    # remove duplicates
    new_df = new_df.drop_duplicates(subset=[batch_h, project_h, prod_date_h, temp_h, ab_stage_h, interval_h])

    # create a merge key for faster matching
    key_cols = [batch_h, project_h, prod_date_h, temp_h, ab_stage_h, interval_h]
    df['_merge_key'] = df[key_cols].astype(str).agg('|'.join, axis=1)
    new_df['_merge_key'] = new_df[key_cols].astype(str).agg('|'.join, axis=1)

    # filter rows with numeric results
    df_numeric = df[df[results_h].apply(lambda x: isinstance(x, float))].copy()
    df_numeric['col_name'] = df_numeric[test_h].astype(str) + ', ' + df_numeric[newUnit_h].astype(str)

    # pivot so we can merge
    df_pivot = df_numeric.pivot_table(index='_merge_key', columns='col_name', values=results_h, aggfunc='first')

    # merge into new_df
    new_df = new_df.merge(df_pivot, left_on='_merge_key', right_index=True, how='left')

    # insert data_type_h column
    new_df.insert(1, data_type_h, 'LIMS Test')

    # drop temporary merge key
    new_df.drop(columns=['_merge_key'], inplace=True)

    return new_df
