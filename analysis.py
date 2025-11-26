import math
import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
from scipy.optimize import curve_fit, OptimizeWarning
from sklearn.metrics import mean_squared_error
import warnings, time

from config_headers import *
from cleaning import fit_columns


def perform_analysis(sheet):

    file_name = sheet.split('/')[-1][len(prefix):]

    print(f"Analyzing {file_name}")
    start_time = time.time()

    # read in re-organized data
    new_df = pd.read_excel(sheet, sheet_name=organized_s, keep_default_na=False)

    # data_df = pd.read_excel(sheet, sheet_name=updated_s, keep_default_na=False)
    # # group formula with nutrient and find average
    # avg_df = find_nut_stats(data_df, new_df)


    print("--Performing Regression Analysis")
    reg_df = compare_regressions(new_df)

    print("--Uploading analysis to the Excel file")
    with pd.ExcelWriter(sheet, mode="a", if_sheet_exists="replace") as writer:
        # avg_df.to_excel(writer, sheet_name=stats_s, index=False)
        # fit_columns(avg_df, writer, stats_s)

        # adding regression analysis
        reg_df.to_excel(writer, sheet_name=regression_s, index=False)
        fit_columns(reg_df, writer, regression_s)

    print("--Computing averages at t=12 months")
    avg12_df = compute_avg12(reg_df)

    print("--Uploading averages to the Excel file")
    with pd.ExcelWriter(sheet, mode="a", if_sheet_exists="replace") as writer:

        # adding regression analysis
        avg12_df.to_excel(writer, sheet_name=average12_s, index=False)
        fit_columns(avg12_df, writer, average12_s)

    # print total time taken
    elapsed = (time.time() - start_time) / 60
    print(f"Finished Analyzing {file_name} in {elapsed:.3f} minutes\n")



# finds average value of each nutrient for all formulas at duration 0
def find_nut_stats(df, new_df):

    # copy relevant columns to new dataframe
    cols = [form_h, temp_h, test_h]
    avg_df = df[cols].copy().drop_duplicates()

    # add new column for average values
    avg_df[avg_h] = pd.Series(dtype=object)

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
            avg_df.loc[i, avg_h] = np.mean(match_nut)
        else:
            drop_rows.append(i)

    return avg_df.drop(drop_rows)


def compare_regressions(df):

    # find column number of test names (starts after interval column)
    interval_index = df.columns.get_loc(interval_h)
    cols = df.columns[interval_index+1:]

    # transpose values under test column to be row data under results column
    df = df.melt(
        id_vars=[form_h, temp_h, interval_h, project_h, run_h, batch_h],
        value_vars=cols,
        var_name=test_h,
        value_name=results_h
    )


    # convert interval and result values to numeric
    df = col_to_num(df, interval_h)
    df = col_to_num(df, results_h)

    # if somehow multiple values for same form, temp, project, run, batch, interval, test combinations, get the average
    agg_df = (
        df.groupby([form_h, temp_h, project_h, run_h, batch_h, interval_h, test_h])[results_h]
        .mean()
        .reset_index()
    )

    # only perform regression if more than 1 data point
    counts = agg_df.groupby([form_h, temp_h, project_h, run_h, batch_h, test_h]).size().reset_index(name="count")

    # Determine if each group has ANY interval >= 60
    interval_check = (
        agg_df.groupby([form_h, temp_h, project_h, run_h, batch_h, test_h])[interval_h]
        .max()
        .reset_index()
    )

    interval_check["has_interval_60plus"] = interval_check[interval_h] >= 60

    # Keep groups with:
    # 1) count > 1
    # 2) at least one interval >= 60
    valid = (
        counts.merge(interval_check, on=[form_h, temp_h, project_h, run_h, batch_h, test_h])
        .query("count > 1 and has_interval_60plus")
        [[form_h, temp_h, project_h, run_h, batch_h, test_h]]
    )
    agg_df = agg_df.merge(valid, on=[form_h, temp_h, project_h, run_h, batch_h, test_h], how="inner")

    # perform regression, rmse using data points for all formula, temp, project, run, batch, test combinations
    final_results = []
    for (formula, temp, project, run, batch, test), group in agg_df.groupby([form_h, temp_h, project_h, run_h, batch_h, test_h]):

        # compute regression as a function of interval months instead of days
        x = group[interval_h].values.astype(float) / 30.0
        y = group[results_h].values.astype(float)

        # --- Linear Regression ---
        slope, intercept = linear_regression(x, y)
        y_linear = slope * x + intercept
        rmse_linear = compute_rmse(y, y_linear)
        formula_linear = f"Result = {slope:.4f}*t + {intercept:.2f}"

        # --- Exponential Regression ---
        a_fit, k_fit, c_fit = 0, 0, 0
        try:
            a_fit, k_fit, c_fit = fitted_regression(x, y)
            y_exp = exp_decay(x, a_fit, k_fit, c_fit)
            rmse_exp = compute_rmse(y, y_exp)
            formula_exp = f"Result = {a_fit:.3f} * e({-k_fit:.3f}*t) + {c_fit:.2f}"
        except:
            rmse_exp = np.inf
            formula_exp = "fit_failed"

        # --- First-order Regression ---
        a_first, k_first = 0, 0
        try:
            a_first, k_first = first_order_regression(x, y)
            y_first = a_first * np.exp(-k_first * x)
            rmse_first = compute_rmse(y, y_first)
            formula_first = f"Result = {a_first:.2f} * exp({-k_first:.4f}*t)"
        except:
            rmse_first = np.inf
            formula_first = "first_order_failed"

        # compare root-mean-square-error of regressions, choose the best
        rmses = {"linear": rmse_linear, "fitted": rmse_exp, "first_order": rmse_first}
        if len(group) == 2:
            best_model = "first_order"
        else:
            best_model = min(rmses, key=rmses.get)

        # Get starting value y0 and estimate y12 + percentage y12/y0
        percent, y0, y12 = get_y0_y12(a_first, a_fit, c_fit, best_model, intercept, k_first, k_fit, slope, x, y)

        # Normalize only the best model
        formula_best = normalize_best(a_first, a_fit, c_fit, best_model, intercept, k_first, k_fit, slope, y0)

        # compile analysis results
        final_results.append({
            form_h: formula,
            project_h: project,
            run_h: run,
            batch_h: batch,
            temp_h: temp,
            test_h: test,
            "Best Model": best_model,
            "Normalized Formula": formula_best,
            t0_h: f"{y0:.5g}",
            t12_h: f"{y12:.5g}",
            percent12_h: f"{percent:.4g}",
            "linear_formula": formula_linear,
            "fitted_formula": formula_exp,
            "first_order": formula_first,
            "linear_rmse": f"{rmse_linear:.3g}",
            "exp_rmse": f"{rmse_exp:.3g}",
            "first_order_rmse": f"{rmse_first:.3g}"
        })

    # convert all analysis results to a dataframe to be uploaded
    final_df = pd.DataFrame(final_results)
    return final_df


# based on best_model, get normalized formula (100% at t=0) of the regression
def normalize_best(a_first, a_fit, c_fit, best_model, intercept, k_first, k_fit, slope, y0):
    formula_best = ''
    if best_model == "linear":
        slope_pct = (slope / y0) * 100
        intercept_pct = (intercept / y0) * 100
        formula_best = f"Percent = {slope_pct:.4g}*t + {intercept_pct:.4g}"

    elif best_model == "fitted":
        A_pct = (a_fit / y0) * 100
        C_pct = (c_fit / y0) * 100
        formula_best = f"Percent = {A_pct:.4g} * exp({-k_fit:.4g}*t) + {C_pct:.4g}"

    elif best_model == "first_order":
        A_first_pct = (a_first / y0) * 100
        formula_best = f"Percent = {A_first_pct:.4g} * exp({-k_first:.4g}*t)"
    return formula_best


# get the value at t=0, estimate t=12M, and get the percentage at t=12M
def get_y0_y12(a_first, a_fit, c_fit, best_model, intercept, k_first, k_fit, slope, x, y):
    if np.any(x == 0):
        y0 = y[x == 0][0]
    else:
        # estimate using the best model at t = 0
        if best_model == "linear":
            y0 = slope * 0 + intercept

        elif best_model == "fitted":
            # exponential regression: A * exp(-k * x) + C
            y0 = a_fit * np.exp(-k_fit * 0) + c_fit

        elif best_model == "first_order":
            # first-order: A * exp(-k * x)
            y0 = a_first * np.exp(-k_first * 0)

    y12 = 0
    # estimate using the best model at t = 12
    if best_model == "linear":
        y12 = slope * 12 + intercept

    elif best_model == "fitted":
        # exponential regression: A * exp(-k * x) + C
        y12 = a_fit * np.exp(-k_fit * 12) + c_fit

    elif best_model == "first_order":
        # first-order: A * exp(-k * x)
        y12 = a_first * np.exp(-k_first * 12)

    if (y0 != 0):
        percent = y12 / y0 * 100
    else:
        percent = math.inf

    return percent, y0, y12


# convert columns to numeric
def col_to_num(df, column):
    df.loc[:, column] = (
        df[column].astype(str)
        .str.strip()
        .replace(r'^\s*$', np.nan, regex=True)
    )
    df.loc[:, column] = pd.to_numeric(df[column], errors='coerce')
    df = df.dropna(subset=[column])
    return df

# root-mean-square error calculation
def compute_rmse(y_true, y_pred):
    return np.sqrt(mean_squared_error(y_true, y_pred))

# perform linear regression analysis
def linear_regression(x, y):
    model = LinearRegression()
    model.fit(x.reshape(-1, 1), y)
    return model.coef_[0], model.intercept_


# for fitted regression
def exp_decay(t, A, k, C):
    return A*np.exp(-k*t) + C

# perform fitted regression analysis
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


# perform first order regression analysis
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


def compute_avg12(df):
    df = col_to_num(df, percent12_h)

    agg_df = (
        df.groupby([form_h, temp_h, test_h])[percent12_h]
        .mean()
        .apply(lambda x: float(f"{x:.4g}"))  # 4 significant figures
        .reset_index()
        .rename(columns={percent12_h: avg12_h})
    )

    return agg_df
