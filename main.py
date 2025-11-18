# This is a sample Python script.
from cleaning import *
from analysis import *
import os
import argparse
import shutil
import traceback

data_folder = "Excel Files"
finshed_folder = "Finished Files"


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', None)

    parser = argparse.ArgumentParser(description="Clean Excel files")
    parser.add_argument(
        "-f", "--file", help="Run on a single Excel file instead of the default folder"
    )
    args = parser.parse_args()

    # secondary unit conversion tab
    unit_df = pd.read_excel(second_conv_file, sheet_name=unit_s, keep_default_na=False, skiprows=3)

    # formula code tab
    form_df = pd.read_excel(formula_code_file, sheet_name=form_s, keep_default_na=False, dtype={project_h: str, run_h: str})

    finished_folder_path = os.path.join(os.path.dirname(__file__), finshed_folder)
    os.makedirs(finished_folder_path, exist_ok=True)

    if args.file:
        # Run on a single file
        file_path = os.path.join(os.path.dirname(__file__), args.file)
        if not os.path.isfile(file_path):
            print(f"Error: {args.file} does not exist.")
        clean_data(unit_df, form_df, file_path)
        # compute_stats(file_path)
        destination = os.path.join(finished_folder_path, os.path.basename(file_path))
        shutil.move(file_path, destination)
    else:
        # Default: run on all Excel files in the "Excel Files" subfolder
        data_folder_path = os.path.join(os.path.dirname(__file__), data_folder)

        if not os.path.exists(data_folder_path):
            print(f"Error: subfolder {data_folder} does not exist.")

        for file in os.listdir(data_folder_path):
            if file.endswith(('.xlsx', '.xls')) and not file.startswith('~'):
                file_path = os.path.join(data_folder_path, file)
                clean_data(unit_df, form_df, file_path)
                # compute_stats(file_path)

                destination = os.path.join(finished_folder_path, os.path.basename(file_path))
                shutil.move(file_path, destination)



# See PyCharm help at https://www.jetbrains.com/help/pycharm/
