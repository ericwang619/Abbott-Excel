import pandas as pd
from datetime import datetime
from decimal import Decimal
from openpyxl.utils import get_column_letter
import time
import re

from config_headers import *

def clean_data(sheet = data_sheet_name):
    file_name = sheet.split('/')[-1]

    print(f"Cleaning {file_name}")
    start_time = time.time()

    # data tab
    data_df = pd.read_excel(sheet, sheet_name=data_s, keep_default_na=False, dtype={project_h: str, run_h: str})

    # unit conversion tab
    unit_df = pd.read_excel(second_conv_file, sheet_name=unit_s, keep_default_na=False, skiprows=3)

    # formula code tab
    form_df = pd.read_excel(formula_code_file, sheet_name=form_s, keep_default_na=False, dtype={project_h: str, run_h: str})


    # add new columns headers to dataframe
    add_columns(data_df, [temp_h, humidity_h, interval_h])

    print("--Converting production date formats")
    data_df[prod_date_h] = data_df[prod_date_h].astype(object)
    data_df[prod_date_h] = data_df[prod_date_h].apply(convert_dates)

    print("--Converting completion date formats")
    data_df[comp_date_h] = data_df[comp_date_h].astype(object)
    data_df[comp_date_h] = data_df[comp_date_h].apply(convert_dates)


    # update column values row by row according to rules
    print("--Converting temperature (C) and humidity")
    data_df[[temp_h, humidity_h]] = data_df.apply(convert_temp_humidity, axis=1)

    print("--Converting intervals to days")
    data_df[interval_h] = data_df[dur_h].apply(convert_duration)


    print("--Adding formula codes and sources")
    data_df = get_formula(data_df, form_df)

    # add_columns(data_df, [form_h, sources_h])
    # for i, row in data_df.iterrows():
    #     if i != 0 and i % 10000 == 0:
    #         print("----processing row", i)
    #
    #     # pull unit conversion and formula values from other sheets
    #     formula, sources = get_formula(row, form_df)
    #     data_df.loc[i, form_h] = formula
    #     data_df.loc[i, sources_h] = sources



    print("--Dropping rows with no formula code")
    # Create a boolean mask for rows to drop
    pattern = '|'.join(re.escape(f) for f in invalid_formulas if f != '')

    # Mask for rows where form_h is empty or contains any invalid formula substring
    mask_invalid = (data_df[form_h] == '') | data_df[form_h].str.contains(pattern, na=False)
    # Keep only the rows that are not invalid
    data_df = data_df[~mask_invalid].reset_index(drop=True)


    print("--Adding test and unit conversions")
    first_conv_df = pd.read_excel(first_conv_file, sheet_name=first_conv_s, keep_default_na=False, skiprows=2)
    data_df = add_unit_conversions(data_df, unit_df, first_conv_df)
    # for i, row in data_df.iterrows():
    #     if i != 0 and i % 10000 == 0:
    #         print("----processing row", i)
    #     test, newUnit, unit_conversion = add_unit_conversions(row, unit_df, first_conv_df)
    #     data_df.loc[i, test_h] = test
    #     data_df.loc[i, newUnit_h] = newUnit
    #     data_df.loc[i, conv_h] = unit_conversion


    data_df[text_h] = data_df[text_h].apply(convert_text)

    add_columns(data_df, [results_h, vitE_h])
    vitE_df = pd.read_excel(vitE_file, sheet_name=vitE_s, keep_default_na=False)
    vitE_map = vitE_df.set_index(form_h)[vitE_h].to_dict()


    print("--Converting to final test results")
    converted = [convert_result(row, vitE_map) for _, row in data_df.iterrows()]

    # --- Unpack into two columns safely ---
    results, vitE_vals = zip(*converted)
    data_df[results_h] = results
    data_df[vitE_h] = vitE_vals

    # for i, row in data_df.iterrows():
    #     if i != 0 and i % 10000 == 0:
    #         print("----processing row", i)
    #     convert_results(data_df, i, vitE_df)


    print("--Removing duplicates")
    data_df = data_df.drop_duplicates(subset=[batch_h, project_h, prod_date_h,
                                              temp_h, ab_stage_h, interval_h, test_h, newUnit_h, results_h])



    print("--Adding updates to the spreadsheet")
    # write revised data to sheet
    with pd.ExcelWriter(sheet, mode="a", if_sheet_exists="replace") as writer:

        data_df.to_excel(writer, sheet_name=updated_s, index=False)
        # fit_columns(data_df, writer, updated_s)

    # consolidate project, batch, temp, duration to have nutrients as columns
    print("--Creating re-organized tab")
    new_df = consolidate(data_df)

    print("--Adding re-organized tab to the spreadsheet")
    # write revised data to sheet
    with pd.ExcelWriter(sheet, mode="a", if_sheet_exists="replace") as writer:

        new_df.to_excel(writer, sheet_name=consolidated_s, index=False)
        fit_columns(new_df, writer, consolidated_s)

    elapsed = (time.time() - start_time) / 60
    print(f"Finished Cleaning {file_name} in {elapsed:.3f} minutes")


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

    # convert production date
    if pd.isna(date):
        parsed_date = pd.NaT
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


# clean up temperature values
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


# convert values to number where applicable
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

# def add_unit_conversions(row, unit_df, first_conv_df):
#     cols = [analysis_h, name_h, unit_h]
#     analysis, name, units = row[cols]
#
#     # if no matching row, default value to empty
#     test = ''
#     newUnit = ''
#     unit_conversion = ''
#
#
#     # check for match in first conversion sheet
#     match = first_conv_df.loc[(first_conv_df[analysis_h] == analysis) &
#                         (first_conv_df[name_h] == name) &
#                         (first_conv_df[unit_h] == units)]
#     if len(match) > 0:
#         test = match.iloc[0][first_conv_test_h]
#         newUnit = match.iloc[0][first_conv_units_h]
#         unit_conversion = match.iloc[0][first_conv_h]
#
#     else:
#
#         # check for matching row in second conversion sheet
#         match = unit_df.loc[(unit_df[analysis_h] == analysis) & (unit_df[name_h] == name) & (unit_df[unit_h] == units)]
#         if len(match) > 0:
#             first_match = match.iloc[0]
#             test = str(first_match[conversion_test_h])
#             newUnit = str(first_match[conversion_units_h])
#             unit_conversion = first_match[conversion_conv_h]
#
#     return test, newUnit, unit_conversion

# --- Helper function for a single row ---
def convert_result(row, vitE_map):
    """
    Convert a row to final result and Vit E value.
    Always returns exactly two values: (result, vitE_val)
    """
    text = row.get(text_h)
    if not isinstance(text, (int, float)):
        return '?', ''  # non-numeric text -> empty result

    result = Decimal(str(text))
    factor = 1
    conversion_factor = row.get(conv_h, '')
    formula = row.get(form_h, '')
    vitE_val = ''

    if conversion_factor not in ['', 'Copy value']:
        terms = str(conversion_factor).split('/')
        # First term
        f0 = terms[0]
        if 'Vit E Factor' in f0:
            vitE_val = vitE_map.get(formula, '')
            if vitE_val in ['', 'Pending', None] or '?' in str(vitE_val):
                return '?', vitE_val
            try:
                factor *= Decimal(str(vitE_val))
            except:
                print('vitE', vitE_val)
                raise Exception
        else:
            if f0[0] == '=':
                f0 = f0[1:]
            try:
                factor *= Decimal(str(float(f0)))
            except:
                return '?', vitE_val
        # Remaining terms
        for t in terms[1:]:
            if t == 'density':
                continue
            elif 'Vit E Factor' in t:
                vitE_val = vitE_map.get(formula, '')
                if vitE_val in ['', 'Pending', None] or '?' in str(vitE_val):
                    return '?', vitE_val
                factor /= Decimal(str(vitE_val))
            else:
                try:
                    factor /= Decimal(str(float(t)))
                except:
                    return '?', vitE_val

    return float(result * factor), vitE_val


# --- Wrapper to return a Series with proper column names ---
def convert_result_series(row, vitE_map):
    try:
        res, vitE_val = convert_result(row, vitE_map)
        return pd.Series([res, vitE_val], index=[results_h, vitE_h])
    except Exception:
        # fallback: always return two values
        return pd.Series(['?', ''], index=[results_h, vitE_h])

# def convert_results(df, i, vitE_df):
#     row = df.loc[i]
#     text = row[text_h]
#     conversion_factor = row[conv_h]
#     if isinstance(text, (int, float)):
#         result = Decimal(str(text))
#     else:
#         df.loc[i, results_h] = '?'
#         return
#     factor = 1
#
#     if conversion_factor != '' and conversion_factor != 'Copy value':
#         terms = str(conversion_factor).split('/')
#         factor = terms[0]
#
#         if 'Vit E Factor' in factor:
#             formula = row[form_h]
#             match = vitE_df.loc[(vitE_df[form_h] == formula)]
#             if len(match) > 0:
#                 vitE = str(match.iloc[0][vitE_h])
#                 df.loc[i, vitE_h] = vitE
#                 if (vitE == 'Pending') or ('?' in vitE):
#                     df.loc[i, results_h] = '?'
#                     return
#
#                 factor = Decimal(str(vitE))
#
#         else:
#             try:
#                 factor = Decimal(str(float(factor)))
#             except:
#                 df.loc[i, results_h] = '?'
#                 return
#
#         for r in range(1, len(terms)):
#             t = terms[r]
#             if t == 'density':
#                 pass
#             elif "Vit E Factor" in t:
#                 formula = row[form_h]
#                 match = vitE_df.loc[(vitE_df[form_h] == formula)]
#                 if len(match) > 0:
#                     vitE = str(match.iloc[0][vitE_h])
#                     df.loc[i, vitE_h] = vitE
#                     if (vitE == 'Pending') or ('?' in vitE):
#                         df.loc[i, results_h] = '?'
#                         return
#                     factor /= Decimal(str(vitE))
#             else:
#                 try:
#                     div = Decimal(str(float(t)))
#                     factor /= div
#                 except:
#                     df.loc[i, results_h] = '?'
#                     return
#     df.loc[i, results_h] = float(result * factor)


# pull matching formula code from formula tab
def get_formula(data_df, form_df):
    # Pick only the relevant columns from lookup
    form_lookup = form_df[[project_h, run_h, batch_h, conversion_formula_h, conversion_sources_h]].copy()

    # Make sure key types match
    for col in [project_h, run_h, batch_h]:
        data_df[col] = data_df[col].astype(object)
        form_lookup[col] = form_lookup[col].astype(object)

    # Merge only the lookup columns (bring them in with default suffixes)
    data_df = data_df.merge(
        form_lookup,
        how='left',
        on=[project_h, run_h, batch_h]
    )

    # After merge, rename the lookup columns to the standard names
    # Use .get() with default to handle the case where they already match
    if conversion_formula_h in data_df.columns:
        data_df.rename(columns={conversion_formula_h: form_h}, inplace=True)
    if conversion_sources_h in data_df.columns:
        data_df.rename(columns={conversion_sources_h: sources_h}, inplace=True)

    # Fill missing values
    data_df[[form_h, sources_h]] = data_df[[form_h, sources_h]].fillna('')

    return data_df

    #
    # project = row[project_h]
    # run = row[run_h]
    # batch = row[batch_h]
    # formula = ''
    # sources = ''
    #
    # # find matching row in other tab if exists
    # match = form_df.loc[(form_df[project_h] == project) &
    #                     (form_df[batch_h] == batch) &
    #                     (form_df[run_h] == run)]
    # if len(match) > 0:
    #     print('found match')
    #     formula = match.iloc[0][conversion_formula_h]
    #     sources = match.iloc[0][conversion_sources_h]
    #
    # return formula, sources


# re-organized sheet
def consolidate(df):

    headers_to_keep = [form_h, project_h, run_h, batch_h, description_h, batch_type_h, batch_sub_h,
                       manu_loc_h, prod_date_h, ab_container_h, ab_stage_h, temp_h, humidity_h, interval_h]
    # copy relevant columns to new dataframe
    new_df = df[headers_to_keep].copy().drop_duplicates()

    # find list of nutrients from newNut tab
    nn_df = pd.read_excel(nutrient_file, sheet_name=nutrient_s, skiprows=1, usecols=[0,1], keep_default_na=False)
    nn_list = set()
    for i, row in nn_df.iterrows():
        nn_list.add(str(row[test_h]) + ', ' + str(row[newUnit_h]))


    # create new column for each nutrient in new dataframe
    for n in nn_list:
        new_df[n] = pd.Series(dtype=object)

    new_df = new_df.drop_duplicates(subset=[batch_h, project_h, prod_date_h,
                                               temp_h, ab_stage_h, interval_h])

    updates = []

    # iterate through all original data, copying over nutrient data to new dataframe
    for i, row in df.iterrows():

        if i != 0 and i % 10000 == 0:
            print('----processing row', i)

        cols = [batch_h, project_h, prod_date_h, temp_h, ab_stage_h, interval_h, test_h, newUnit_h, results_h]
        batch, project, production_date, temp, ab_stage, interval, test, units, results = row[cols]
        if results == '' or not isinstance(results, float):
            continue

        # find matching row in new dataframe
        match = new_df.loc[(new_df[batch_h] == batch) &
                           (new_df[project_h] == project) &
                           (new_df[prod_date_h] == production_date) &
                           (new_df[temp_h] == temp) &
                           (new_df[ab_stage_h] == ab_stage) &
                           (new_df[interval_h] == interval)]

        # copy over nutrient value to corresponding column
        if len(match) > 0:
            index = match.index[0]
            col_name = str(test) + (', ' + str(units) if units else '')
            updates.append((index, col_name, float(results)))

    # apply all at once
    for index, col_name, value in updates:
        new_df.at[index, col_name] = value

    new_df.insert(1, data_type_h, 'LIMS Test')

    return new_df

