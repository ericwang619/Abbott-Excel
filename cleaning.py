import pandas as pd
from datetime import datetime
from decimal import Decimal
from openpyxl.utils import get_column_letter
import time

from config_headers import *

def clean_data(sheet = data_sheet_name):
    file_name = sheet.split('/')[-1]

    print(f"Cleaning {file_name}")
    start_time = time.time()

    # data tab
    data_df = pd.read_excel(sheet, sheet_name=data_s, keep_default_na=False)

    # unit conversion tab
    unit_df = pd.read_excel(second_conv_file, sheet_name=unit_s, keep_default_na=False, skiprows=3)

    # formula code tab
    form_df = pd.read_excel(formula_code_file, sheet_name=form_s, keep_default_na=False)


    # add new columns headers to dataframe
    add_columns(data_df)

    print("-Converting production date formats")
    data_df[prod_date_h] = data_df[prod_date_h].astype(object)
    for i, _ in data_df.iterrows():
        if i != 0 and i % 10000 == 0:
            print("--processing row", i)
        prod_date = convert_dates(data_df, i, prod_date_h)
        data_df.loc[i, prod_date_h] = prod_date

    print("-Converting completion date formats")
    data_df[comp_date_h] = data_df[comp_date_h].astype(object)
    for i, _ in data_df.iterrows():
        if i != 0 and i % 10000 == 0:
            print("--processing row", i)
        comp_date = convert_dates(data_df, i, comp_date_h)
        data_df.loc[i, comp_date_h] = comp_date


    # update column values row by row according to rules
    print("-Converting temperature (C) and humidity")
    for i, _ in data_df.iterrows():
        if i != 0 and i % 10000 == 0:
            print("--processing row", i)

        # cleanup data for temperature and humidity
        (temp, humidity) = convert_temp_humidity(data_df, i)
        data_df.loc[i, temp_h] = temp
        data_df.loc[i, humidity_h] = humidity

    print("-Converting intervals to days")
    for i, _ in data_df.iterrows():
        # duration conversion
        data_df.loc[i, interval_h] = convert_duration(data_df, i)

    print("-Adding formula codes and sources")
    no_formula_rows = []
    for i, _ in data_df.iterrows():
        if i != 0 and i % 10000 == 0:
            print("--processing row", i)

        # pull unit conversion and formula values from other sheets
        formula, sources = get_formula(data_df, i, form_df)
        if formula == '' or any(f in str(formula) for f in invalid_formulas):
            no_formula_rows.append(i)
        else:
            data_df.loc[i, form_h] = formula
            data_df.loc[i, sources_h] = sources

    print("-Dropping rows with no formula code")
    data_df.drop(no_formula_rows, inplace=True)
    data_df = data_df.reset_index(drop=True)

    print("-Adding test and unit conversions")
    first_conv_df = pd.read_excel(first_conv_file, sheet_name=first_conv_s, keep_default_na=False, skiprows=2)
    for i, _ in data_df.iterrows():
        if i != 0 and i % 10000 == 0:
            print("--processing row", i)
        add_unit_conversions(data_df, i, unit_df, first_conv_df)


    vitE_df = pd.read_excel(vitE_file, sheet_name=vitE_s, keep_default_na=False)
    print("-Converting to final test results")
    for i, _ in data_df.iterrows():
        if i != 0 and i % 10000 == 0:
            print("--processing row", i)
        # convert texts to float values if possible
        data_df.loc[i, text_h] = convert_text(data_df, i)
        convert_results(data_df, i, vitE_df)


    print("-Removing duplicates")
    data_df = data_df.drop_duplicates(subset=[batch_h, project_h, prod_date_h,
                                              temp_h, ab_stage_h, interval_h, test_h, newUnit_h, results_h])

    # consolidate project, batch, temp, duration to have nutrients as columns
    print("-Creating re-organized sheet")
    new_df = consolidate(data_df)

    print("-Adding updates to the spreadsheet")
    # write revised data to sheet
    with pd.ExcelWriter(sheet, mode="a", if_sheet_exists="replace") as writer:

        data_df.to_excel(writer, sheet_name=updated_s, index=False)
        fit_columns(data_df, writer, updated_s)

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
def add_columns(df):
    for h in new_headers:
        df[h] = pd.Series(dtype=object)


# convert production and completion date formats
def convert_dates(df, i, col):

    row = df.loc[i]

    # convert production date
    date = row[col]
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
def convert_temp_humidity(df, i):
    row = df.loc[i]
    storage = str(row[storage_h])
    humidity = ''
    if storage == 'ROOM':
        return 22, humidity
    elif storage == 'REFRIG':
        return 4, humidity
    elif storage == "FROZEN":
        return -20, humidity
    elif 'C' in storage:
        # strip 'C' from number
        i = storage.index('C')
        temp = int(storage[:i])
        humidity = storage[i+2:]
        if humidity != '':
            humidity = int(humidity)
        return temp, humidity
    elif 'F' in storage:
        # calculate C from F value
        i = storage.index('F')
        temp = round((int(storage[:i])-32)/9*5, 2)
        humidity = storage[i+2:]
        if humidity != '':
            humidity = int(humidity)
        return temp, humidity
    return storage, humidity


# convert duration values to # days
def convert_duration(df, i):
    row = df.loc[i]
    dur = str(row[dur_h])
    if 'D' in dur:
        return int(dur[:dur.index('D')])
    elif 'M' in dur:
        return 30*int(dur[:dur.index('M')])
    return dur


# convert values to number where applicable
def convert_text(df, i):
    row = df.loc[i]
    text = str(row[text_h])
    try:
        text = float(text)
    finally:
        return text


# pull matching columns from conversion tab
def add_unit_conversions(df, i, unit_df, first_conv_df):
    row = df.loc[i]
    cols = [analysis_h, name_h, unit_h]
    analysis, name, units = row[cols]

    # if no matching row, default value to empty
    test = ''
    newUnit = ''
    unit_conversion = ''


    # check for match in first conversion sheet
    match = first_conv_df.loc[(first_conv_df[analysis_h] == analysis) &
                        (first_conv_df[name_h] == name) &
                        (first_conv_df[unit_h] == units)]
    if len(match) > 0:
        test = match.iloc[0][first_conv_test_h]
        newUnit = match.iloc[0][first_conv_units_h]
        unit_conversion = match.iloc[0][first_conv_h]

    else:

        # check for matching row in second conversion sheet
        match = unit_df.loc[(unit_df[analysis_h] == analysis) & (unit_df[name_h] == name) & (unit_df[unit_h] == units)]
        if len(match) > 0:
            first_match = match.iloc[0]
            test = str(first_match[conversion_test_h])
            newUnit = str(first_match[conversion_units_h])
            unit_conversion = first_match[conversion_conv_h]

    # add values to new columns
    df.loc[i, test_h] = test
    df.loc[i, newUnit_h] = newUnit
    df.loc[i, conv_h] = unit_conversion

def convert_results(df, i, vitE_df):
    row = df.loc[i]
    text = row[text_h]
    conversion_factor = row[conv_h]
    if isinstance(text, (int, float)):
        result = Decimal(str(text))
    else:
        df.loc[i, results_h] = '?'
        return
    factor = 1

    if conversion_factor != '' and conversion_factor != 'Copy value':
        terms = str(conversion_factor).split('/')
        factor = terms[0]

        if 'Vit E Factor' in factor:
            formula = row[form_h]
            match = vitE_df.loc[(vitE_df[form_h] == formula)]
            if len(match) > 0:
                vitE = match.iloc[0][vitE_h]
                df.loc[i, vitE_h] = vitE
                if (vitE == 'Pending') or ('?' in vitE):
                    df.loc[i, results_h] = '?'
                    return

                factor = Decimal(str(vitE))

        else:
            try:
                factor = Decimal(str(float(factor)))
            except:
                df.loc[i, results_h] = '?'
                return


        for r in range(1, len(terms)):
            t = terms[r]
            if t == 'density':
                pass
            elif "Vit E Factor" in t:
                formula = row[form_h]
                match = vitE_df.loc[(vitE_df[form_h] == formula)]
                if len(match) > 0:
                    vitE = match.iloc[0][vitE_h]
                    df.loc[i, vitE_h] = vitE
                    if (vitE == 'Pending') or ('?' in vitE):
                        df.loc[i, results_h] = '?'
                        return
                    factor /= Decimal(str(vitE))
            else:
                try:
                    div = Decimal(str(float(t)))
                    factor /= div
                except:
                    df.loc[i, results_h] = '?'
                    return
    df.loc[i, results_h] = float(result * factor)


# pull matching formula code from formula tab
def get_formula(df, i, form_df):
    row = df.loc[i]
    project = row[project_h]
    run = row[run_h]
    batch = row[batch_h]
    formula = ''
    sources = ''

    # find matching row in other tab if exists
    match = form_df.loc[(form_df[project_h] == project) &
                        (form_df[batch_h] == batch) &
                        (form_df[run_h] == run)]
    if len(match) > 0:
        formula = match.iloc[0][conversion_formula_h]
        sources = match.iloc[0][conversion_sources_h]

    return formula, sources


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


    # iterate through all original data, copying over nutrient data to new dataframe
    for i, row in df.iterrows():

        if i != 0 and i % 10000 == 0:
            print('--processing row', i)

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
            new_df.loc[index, col_name] = float(results)

    new_df.insert(1, data_type_h, 'LIMS Test')

    return new_df

