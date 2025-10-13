# Abbott-Excel

This program uses python to clean up Abbott spreadsheet data. It will clean up and reorganize data within all excel files inside the "Excel Files" subfolder

## For your first time:
Please download python onto your computer before using (https://www.python.org/downloads/)
1. Click "Download Python 3.x.x", using the latest stable version
2. Run the installer, remember to check "Add Python 3.x to PATH" at the bottom, then "Install Now"
3. Verify python installation. Open command prompt and type "python --version"

You will also need to create a virtual environment and install the pandas module into it. This will prevent this program from affecting anything else on your computer:
1. Create and navigate to the folder that you will use for this program
2. type "python -m venv venv" to create your virtual environment
3. type "venv\Scripts\activate" to activate your virtual environment
4. type "python -m pip install --upgrade pip"
5. type "pip install pandas"


## To use the program:
1. download the main.py, cleaning.py, and config_headers.py files into your project folder
2. Create a subfolder called "Excel Files" and store any spreadsheets you want to edit inside
3. Modify the config_headers.py file to match your tab and column names
4. open up the command line, navigate to the project folder
5. Activate your virtual environment with "venv\Scripts\activate" if not already activated.
6. type in "python main.py" and hit enter to run the program
7. Once you are done executing the program, type "Deactivate" to close your virtual environment