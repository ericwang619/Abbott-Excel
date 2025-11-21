import math

import pandas as pd
from decimal import Decimal
import numpy as np
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
import matplotlib.pyplot as plt
from sklearn.decomposition import PCA
import seaborn as sns
from sklearn.linear_model import LinearRegression
from scipy.optimize import curve_fit, OptimizeWarning
from sklearn.metrics import mean_squared_error
import warnings, time

from config_headers import *
from cleaning import fit_columns


def perform_analysis(sheet = data_sheet_name):

    file_name = sheet.split('/')[-1]

    print(f"Analyzing {file_name}")
    start_time = time.time()
    # data_df = pd.read_excel(sheet, sheet_name=updated_s, keep_default_na=False)
    new_df = pd.read_excel(sheet, sheet_name=organized_s, keep_default_na=False)

    # # group formula with nutrient and find average
    # avg_df = find_nut_stats(data_df, new_df)
    #
    # # sort formula into clusters
    # cluster_df = create_cluster(avg_df)

    print("--Performing Regression Analysis")
    reg_df = compare_regressions(new_df)

    print("--Adding updates to the spreadsheet")
    with pd.ExcelWriter(sheet, mode="a", if_sheet_exists="replace") as writer:
        # avg_df.to_excel(writer, sheet_name=stats_s, index=False)
        # cluster_df.to_excel(writer, sheet_name=clusters_s, index=False)
        #
        # fit_columns(avg_df, writer, stats_s)
        # fit_columns(cluster_df, writer, clusters_s)

        reg_df.to_excel(writer, sheet_name=regression_s, index=False)
        fit_columns(reg_df, writer, regression_s)

    # print total time taken
    elapsed = (time.time() - start_time) / 60
    print(f"Finished Analyzing {file_name} in {elapsed:.3f} minutes\n")



# finds average value of each nutrient for all formulas at duration 0
def find_nut_stats(df, new_df):

    # copy relevant columns to new dataframe
    cols = [form_h, storage_h, test_h]
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
                           (new_df[storage_h] == temp) &
                           new_df[nut]]
        if len(match) > 0:
            match_nut = match[nut]
            avg_df.loc[i, avg_h] = np.mean(match_nut)
            avg_df.loc[i, min_h] = np.min(match_nut)
            avg_df.loc[i, max_h] = np.max(match_nut)
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


def compare_regressions(df):

    interval_index = df.columns.get_loc(interval_h)
    cols = df.columns[interval_index+1:]

    df = df.melt(
        id_vars=[form_h, temp_h, interval_h, project_h, run_h, batch_h],
        value_vars=cols,
        var_name=test_h,
        value_name=results_h
    )


    # Clean interval column robustly
    df = col_to_num(df, interval_h)
    df = col_to_num(df, results_h)

    agg_df = (
        df.groupby([form_h, temp_h, project_h, run_h, batch_h, interval_h, test_h])[results_h]
        .mean()
        .reset_index()
    )

    counts = agg_df.groupby([form_h, temp_h, project_h, run_h, batch_h, test_h]).size().reset_index(name="count")
    valid = counts[counts["count"]>1][[form_h, temp_h, project_h, run_h, batch_h, test_h]]
    agg_df = agg_df.merge(valid, on=[form_h, temp_h, project_h, run_h, batch_h, test_h], how="inner")

    final_results = []
    for (formula, temp, project, run, batch, test), group in agg_df.groupby([form_h, temp_h, project_h, run_h, batch_h, test_h]):
        x = group[interval_h].values.astype(float) / 30.0
        y = group[results_h].values.astype(float)

        slope, intercept = linear_regression(x, y)
        y_linear = slope * x + intercept
        rmse_linear = compute_rmse(y, y_linear)
        formula_linear = f"Result = {slope:.4f}*t + {intercept:.2f}"

        # --- Exponential regression ---
        try:
            A_fit, k_fit, C_fit = fitted_regression(x, y)
            y_exp = exp_decay(x, A_fit, k_fit, C_fit)
            rmse_exp = compute_rmse(y, y_exp)
            formula_exp = f"Result = {A_fit:.3f} * e({-k_fit:.3f}*t) + {C_fit:.2f}"
        except:
            rmse_exp = np.inf
            formula_exp = "fit_failed"

        # --- First-order regression ---
        try:
            A_first, k_first = first_order_regression(x, y)
            y_first = A_first * np.exp(-k_first * x)
            rmse_first = compute_rmse(y, y_first)
            formula_first = f"Result = {A_first:.2f} * exp({-k_first:.4f}*t)"
        except:
            rmse_first = np.inf
            formula_first = "first_order_failed"


        rmses = {"linear": rmse_linear, "fitted": rmse_exp, "first_order": rmse_first}
        best_model = min(rmses, key=rmses.get)

        # Get starting value y0 and estimate y12

        if np.any(x == 0):
            y0 = y[x == 0][0]
        else:
            # estimate using the best model at t = 0
            if best_model == "linear":
                y0 = slope * 0 + intercept

            elif best_model == "fitted":
                # exponential regression: A * exp(-k * x) + C
                y0 = A_fit * np.exp(-k_fit * 0) + C_fit

            elif best_model == "first_order":
                # first-order: A * exp(-k * x)
                y0 = A_first * np.exp(-k_first * 0)

        y12 = 0
        # estimate using the best model at t = 12
        if best_model == "linear":
            y12 = slope * 12 + intercept

        elif best_model == "fitted":
            # exponential regression: A * exp(-k * x) + C
            y12 = A_fit * np.exp(-k_fit * 12) + C_fit

        elif best_model == "first_order":
            # first-order: A * exp(-k * x)
            y12 = A_first * np.exp(-k_first * 12)

        # Normalize only the best model
        formula_best = ''
        if best_model == "linear":
            slope_pct = (slope / y0) * 100
            intercept_pct = (intercept / y0) * 100
            formula_best = f"Percent = {slope_pct:.4g}*t + {intercept_pct:.4g}"

        elif best_model == "fitted":
            A_pct = (A_fit / y0) * 100
            C_pct = (C_fit / y0) * 100
            formula_best = f"Percent = {A_pct:.4g} * exp({-k_fit:.4g}*t) + {C_pct:.4g}"

        elif best_model == "first_order":
            A_first_pct = (A_first / y0) * 100
            formula_best = f"Percent = {A_first_pct:.4g} * exp({-k_first:.4g}*t)"

        if (y0 != 0):
            percent = y12/y0 * 100
        else:
            percent = math.inf

        final_results.append({
            form_h: formula,
            project_h: project,
            run_h: run,
            batch_h: batch,
            temp_h: temp,
            test_h: test,
            "Best Model": best_model,
            "Normalized Formula": formula_best,
            "Result at t=0M": f"{y0:.5g}",
            "Result at t=12M": f"{y12:.5g}",
            "% Remaining": f"{percent:.4g}%",
            "linear_formula": formula_linear,
            "fitted_formula": formula_exp,
            "first_order": formula_first,
            "linear_rmse": f"{rmse_linear:.3g}",
            "exp_rmse": f"{rmse_exp:.3g}",
            "first_order_rmse": f"{rmse_first:.3g}"
        })

    final_df = pd.DataFrame(final_results)
    return final_df


def col_to_num(df, column):
    df.loc[:, column] = (
        df[column].astype(str)
        .str.strip()
        .replace(r'^\s*$', np.nan, regex=True)
    )
    df.loc[:, column] = pd.to_numeric(df[column], errors='coerce')
    df = df.dropna(subset=[column])
    return df

def compute_rmse(y_true, y_pred):
    return np.sqrt(mean_squared_error(y_true, y_pred))

def linear_regression(x, y):
    model = LinearRegression()
    model.fit(x.reshape(-1, 1), y)
    return model.coef_[0], model.intercept_

def exp_decay(t, A, k, C):
    return A*np.exp(-k*t) + C

def fitted_regression(x, y):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", OptimizeWarning)
        warnings.simplefilter("ignore", RuntimeWarning)
        # Fit exponential decay: y = A * exp(-k*t) + C
        params, _ = curve_fit(
            lambda t, A, k, C: A * np.exp(-k * t) + C,
            x, y,
            p0=(y[0], 0.01, 0)
        )
        A, k, C = params

    return A, k, C


def first_order_regression(x, y):
    # Must have strictly positive y-values for log() to work
    if np.any(y <= 0) or np.any(np.isnan(y)):
        raise ValueError("First-order regression invalid due to non-positive or NaN y-values.")

    y_log = np.log(y)

    model = LinearRegression()
    model.fit(x.reshape(-1, 1), y_log)

    k = -model.coef_[0]
    A = np.exp(model.intercept_)

    return A, k