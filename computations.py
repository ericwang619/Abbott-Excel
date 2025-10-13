import pandas as pd
from decimal import Decimal
import statistics
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
import matplotlib.pyplot as plt
from sklearn.decomposition import PCA
import seaborn as sns

from config_headers import *


def compute_stats(sheet = sheet_name):
    data_df = pd.read_excel(sheet, sheet_name=data_s, keep_default_na=False)
    new_df = pd.read_excel(sheet, sheet_name=consolidated_s, keep_default_na=False)

    # group formula with nutrient and find average
    avg_df = find_nut_stats(data_df, new_df)
    with pd.ExcelWriter(sheet, mode="a", if_sheet_exists="replace") as writer:
        avg_df.to_excel(writer, sheet_name=stats_s, index=False)

    cluster_df = create_cluster(avg_df)
    with pd.ExcelWriter(sheet, mode="a", if_sheet_exists="replace") as writer:
        cluster_df.to_excel(writer, sheet_name=clusters_s, index=False)



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