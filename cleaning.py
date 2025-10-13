import pandas as pd
from decimal import Decimal
import statistics
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
import matplotlib.pyplot as plt
from sklearn.decomposition import PCA
import seaborn as sns

sheet_name = 'Data Sample for programming.xlsx'

# sheet names
data_s = 'Data'
unit_s = 'Unit Conversion'
form_s = 'FormulaList'
newNut_s = 'NewNut'

# data headers
batch_h = 'BATCH'
project_h = 'PROJECT'
run_h = 'RUN'
prod_h = 'PRODUCTION'
temp_h = 'Temperature'
dur_h = 'Duration'
method_h = 'Method'
nut_h = 'Nut'
text_h = 'TEXT'
unit_h = 'UNITS'

# new headers
conv_h = 'Unit Conversion Factor'
newNut_h = 'NewNut'
newUnit_h = 'New Unit'
adjv_h = 'Adjusted Value'
form_h = 'Formula'
avg_h = 'Average'
min_h = 'Minimum'
max_h = 'Maximum'
count_h = 'Count'

def clean_data():
    data_df = pd.read_excel(sheet_name, sheet_name=data_s, keep_default_na=False)
    conv_df = pd.read_excel(sheet_name, sheet_name=unit_s, keep_default_na=False)
    form_df = pd.read_excel(sheet_name, sheet_name=form_s, keep_default_na=False)

    # remove duplicates
    data_df = data_df.drop_duplicates(subset=[batch_h, project_h, temp_h,
                                    dur_h, method_h, nut_h, text_h, unit_h])

    # add new columns headers to dataframe
    add_columns(data_df)

    # update column values row by row according to rules
    for i, _ in data_df.iterrows():

        # cleanup data for temperature, duration, text
        data_df.loc[i, temp_h] = convert_temp(data_df, i)
        data_df.loc[i, dur_h] = convert_duration(data_df, i)
        data_df.loc[i, text_h] = convert_text(data_df, i)

        # pull unit conversion and formula values from other sheets
        add_conversions(data_df, i, conv_df)
        add_formula(data_df, i, form_df)

    # write revised data to sheet
    with pd.ExcelWriter(sheet_name, mode="a", if_sheet_exists="replace") as writer:
        data_df.to_excel(writer, sheet_name="UpdatedData", index=False)

    # consolidate project, batch, temp, duration with nutrients
    new_df = consolidate(data_df)
    with pd.ExcelWriter(sheet_name, mode="a", if_sheet_exists="replace") as writer:
        new_df.to_excel(writer, sheet_name="ConsolidatedData", index=False)

    # group formula with nutrient and find average
    avg_df = find_nut_stats(data_df, new_df)
    with pd.ExcelWriter(sheet_name, mode="a", if_sheet_exists="replace") as writer:
        avg_df.to_excel(writer, sheet_name="FormulaStats", index=False)

    cluster_df = create_cluster(avg_df)
    with pd.ExcelWriter(sheet_name, mode="a", if_sheet_exists="replace") as writer:
        cluster_df.to_excel(writer, sheet_name="Clusters", index=False)



# adds new column headers to dataframe
def add_columns(df):
    df[conv_h] = pd.Series(dtype=object)
    df[newNut_h] = pd.Series(dtype=object)
    df[newUnit_h] = pd.Series(dtype=object)
    df[adjv_h] = pd.Series(dtype=object)
    df[form_h] = pd.Series(dtype=object)


# clean up temperature values
def convert_temp(df, i):
    row = df.loc[i]
    temp = str(row[temp_h])
    if temp == 'ROOM':
        return 22
    elif temp == 'REFRIG':
        return 4
    elif 'C' in temp:
        # strip 'C' from number
        i = temp.index('C')
        return int(temp[:i])
    elif 'F' in temp:
        # calculate C from F value
        i = temp.index('F')
        return round((int(temp[:i])-32)/9*5, 2)
    return temp


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
def add_conversions(df, i, conv_df):
    row = df.loc[i]
    cols = [method_h, nut_h, unit_h]
    method, nut, units = row[cols]

    # if no matching row, default value to n/a
    newNut = 'n/a'
    newUnit = 'n/a'
    unit_conversion = 'n/a'
    adjusted_value = 'n/a'

    # find matching row in conversion tab if it exists
    match = conv_df.loc[(conv_df[method_h] == method) & (conv_df[nut_h] == nut) & (conv_df[unit_h] == units)]
    if len(match) > 0:
        first_match = match.iloc[0]
        newNut = str(first_match[newNut_h])
        newUnit = str(first_match[newUnit_h])
        unit_conversion = first_match[conv_h]

        # perform arithmetic conversion
        if (isinstance(row[text_h], float) & (isinstance(unit_conversion, float) | isinstance(unit_conversion, int))):
            text = Decimal(str(row[text_h]))
            conversion = Decimal(str(unit_conversion))
            adjusted_value = float(text * conversion)

    # add values to new columns
    df.loc[i, newNut_h] = newNut
    df.loc[i, newUnit_h] = newUnit
    df.loc[i, conv_h] = unit_conversion
    df.loc[i, adjv_h] = adjusted_value


# pull matching formula code from formula tab
def add_formula(df, i, form_df):
    row = df.loc[i]
    project = row[project_h]
    batch = row[batch_h]
    formula = 'n/a'

    # find matching row in other tab if exists
    match = form_df.loc[(form_df[project_h] == project) & (form_df[batch_h] == batch)]
    if len(match) > 0:
        formula = match.iloc[0][form_h]

    df.loc[i, form_h] = formula


# creates new sheet with project, batch, temperature, duration, formula + all newNuts as column headers
def consolidate(df):

    # copy relevant columns to new dataframe
    cols_new = [project_h, batch_h, temp_h, dur_h, form_h]
    new_df = df[cols_new].copy().drop_duplicates()

    # find list of nutrients from newNut tab
    nn_df = pd.read_excel(sheet_name, sheet_name=newNut_s, header=None, names=[newNut_h])
    nn_list = list(set(nn_df[newNut_h]))

    # create new column for each nutrient in new dataframe
    for n in nn_list:
        new_df[n] = pd.Series(dtype=object)

    # iterate through all original data, copying over nutrient data to new dataframe
    for i, row in df.iterrows():
        cols = [project_h, batch_h, temp_h, dur_h, form_h, newNut_h, adjv_h]
        project, batch, temp, dur, form, nut, value = row[cols]
        if (value == 'n/a'):
            continue

        # find matching row in new dataframe
        match = new_df.loc[(new_df[project_h] == project) &
                           (new_df[batch_h] == batch) &
                           (new_df[temp_h] == temp) &
                           (new_df[dur_h] == dur) &
                           (new_df[form_h] == form)]

        # copy over nutrient value to corresponding column
        if len(match) > 0:
            index = match.index[0]
            new_df.loc[index, nut] = float(value)

    return new_df


# finds average value of each nutrient for all formulas at duration 0
def find_nut_stats(df, new_df):

    # copy relevant columns to new dataframe
    cols = [form_h, temp_h, newNut_h]
    avg_df = df[cols].copy().drop_duplicates()

    # add new column for average values
    avg_df[avg_h] = pd.Series(dtype=object)
    avg_df[min_h] = pd.Series(dtype=object)
    avg_df[max_h] = pd.Series(dtype=object)
    avg_df[count_h] = pd.Series(dtype=object)

    drop_rows = []

    for i, row in avg_df.iterrows():
        form, temp, nut = row[cols]

        if (form == 'n/a') | (nut == 'n/a'):
            drop_rows.append(i)
            continue

        match = new_df.loc[((new_df[dur_h] == 0) | (new_df[dur_h] == 'n/a')) &
                           (new_df[form_h] == form) &
                           (new_df[temp_h] == temp) &
                           new_df[nut]]
        if len(match) > 0:
            match_nut = match[nut]
            avg_df.loc[i, avg_h] = statistics.mean(match_nut)
            avg_df.loc[i, min_h] = min(match_nut)
            avg_df.loc[i, max_h] = max(match_nut)
            avg_df.loc[i, count_h] = len(match)
        else:
            drop_rows.append(i)

    return avg_df.drop(drop_rows)


def create_cluster(avg_df):
    cluster_df = avg_df.pivot(index='Formula', columns='NewNut', values='Average')
    cluster_df = cluster_df.fillna(0).infer_objects(copy=False)

    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(cluster_df)

    kmeans = KMeans(n_clusters=5, random_state=42)
    clusters = kmeans.fit_predict(X_scaled)
    cluster_df = cluster_df.reset_index()
    formula_col = cluster_df.pop('Formula')
    cluster_df.insert(0, 'Formula', formula_col)
    cluster_df.insert(0, 'Cluster', clusters)

    return cluster_df.sort_values(by='Cluster')