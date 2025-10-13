# This is a sample Python script.
from cleaning import *
import os
import argparse

folder_name = "Excel Files"


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', None)

    parser = argparse.ArgumentParser(description="Clean Excel files")
    parser.add_argument(
        "-f", "--file", help="Run on a single Excel file instead of the default folder"
    )
    args = parser.parse_args()

    if args.file:
        # Run on a single file
        file_path = os.path.join(os.path.dirname(__file__), args.file)
        if not os.path.isfile(file_path):
            print(f"Error: {args.file} does not exist.")
        clean_data(file_path)
    else:
        # Default: run on all Excel files in the "Excel Files" subfolder
        folder_path = os.path.join(os.path.dirname(__file__), folder_name)
        if not os.path.exists(folder_path):
            print(f"Error: subfolder {folder_name} does not exist.")

        for file in os.listdir(folder_path):
            if file.endswith(('.xlsx', '.xls')):
                file_path = os.path.join(folder_path, file)
                clean_data(file_path)


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
