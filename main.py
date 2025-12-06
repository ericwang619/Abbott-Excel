# This is a sample Python script.
from cleaning import *
from analysis import *
import os
import argparse
import shutil




# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', None)

    # determine if clean (default), analyze, or both
    parser = argparse.ArgumentParser(description="Clean Excel files")
    parser.add_argument(
        "-a", "--analyze",
        action="store_true",
        help="Run perform_analysis() only, skip clean_data()"
    )
    parser.add_argument(
        "-b", "--both",
        action="store_true",
        help="Run both clean_data() and perform_analysis()"
    )
    args = parser.parse_args()

    run_clean = True
    run_analysis = True

    if args.both:
        run_clean = True
        run_analysis = True
    elif args.analyze:
        run_clean = False
        run_analysis = True

    # store secondary unit conversion tab
    unit_df = pd.read_excel(Path(helper_folder) / second_conv_file, sheet_name=unit_s, keep_default_na=False, na_values=[], skiprows=3)
    # store formula code tab
    form_df = pd.read_excel(Path(helper_folder) / formula_code_file, sheet_name=form_s, keep_default_na=False, na_values=[], dtype={project_h: str, run_h: str})

    # make a subfolder to store updated files
    finished_folder_path = os.path.join(os.path.dirname(__file__), finshed_folder)
    os.makedirs(finished_folder_path, exist_ok=True)

    # check if subfolder with original data files exists
    data_folder_path = os.path.join(os.path.dirname(__file__), data_folder)
    if not os.path.exists(data_folder_path):
        print(f"Error: subfolder {data_folder} does not exist.")

    # run program on all files inside subfolder
    for file in os.listdir(data_folder_path):
        if file.endswith(('.xlsx', '.xls')) and not file.startswith('~'):

            file_path = os.path.join(data_folder_path, file)

            # create new file to store updates
            new_file = prefix + file
            new_file_path = os.path.join(data_folder_path, new_file)
            shutil.copy(file_path, new_file_path)

            try:

                # Run your operations on the new file
                if run_clean:
                    clean_data(unit_df, form_df, new_file_path)

                if run_analysis:
                    perform_analysis(new_file_path)

                # Move final file to finished folder
                destination = os.path.join(finished_folder_path, new_file)
                if os.path.exists(destination):
                    os.remove(destination)  # replace if file name exists already
                shutil.move(new_file_path, destination)

                # Extract each sheet and save as its own file
                print("--Extracting tabs to individual sheets")
                all_sheets = pd.read_excel(destination, sheet_name=None, keep_default_na=False)
                for sheet_name, df in all_sheets.items():
                    # Build output path: finished_folder_path / new_file_tabName.xlsx
                    name_only = os.path.splitext(new_file)[0]  # remove .xlsx
                    sheet_file = os.path.join(finished_folder_path, f"{name_only}_{sheet_name}.xlsx")

                    # Save the sheet
                    df.to_excel(sheet_file, index=False)

            except Exception as e:
                print(f"Error processing {file}: {e}")

                # Delete the new file if something went wrong
                if os.path.exists(new_file_path):
                    try:
                        os.remove(new_file_path)
                    except Exception as delete_err:
                        print(f"Failed to delete {new_file}: {delete_err}")



# See PyCharm help at https://www.jetbrains.com/help/pycharm/
