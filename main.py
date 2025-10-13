# This is a sample Python script.
from cleaning import *
import os

folder_name = "Excel Files"


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', None)

    for file in os.listdir(folder_name):
        if file.endswith('.xlsx'):
            clean_data(os.path.join(folder_name, file))


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
