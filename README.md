# Abbott-Excel

This program uses python to clean up Abbott spreadsheet data.  
It will clean up and reorganize data within all excel files inside the "Excel Files" subfolder.  
Functionality for statistics and projections will be added later on. 

## For your first time:
Please download python onto your computer before using (https://www.python.org/downloads/)
1. Click "Download Python 3.x.x", using the latest stable version
2. Run the installer, remember to check "Add Python.exe to PATH" at the bottom, then "Install Now"
3. Verify python installation. Open command prompt and type "python --version"

You will also need to create a virtual environment and install the pandas, openpyxl modules into it.  
This will prevent this program from affecting anything else on your computer:
1. Create and navigate to the folder that you will use for this program.  
Right click the folder in file explorer, and click "copy as path"  
Open command prompt from the start menu, type "cd " and paste the file path
2. Type "python -m venv venv" to create your virtual environment
3. Type "venv\Scripts\activate" to activate your virtual environment
4. Type "python -m pip install --upgrade pip"
5. Type "pip install pandas"
6. Type "pip install openpyxl"
7. Type "deactivate" to deactivate the virtual environment


## To use the program:
1. Download the main.py, cleaning.py, and config_headers.py files from the github repository into your project folder
2. Create a subfolder called "Excel Files" within your project folder and store any spreadsheets you want to edit inside
3. Copy into the project folder the files that will be utilized for cleaning (ex. formula codes, nutrient list, etc.)  
   Note: these same files will be used for all data files within the Excel Files subfolder
3. Modify the config_headers.py file with a text editor to match your file, tab, and column header names. Note that the program will use the same formatting for all spreadsheets. 
4. Open up the command line, navigate to the project folder
   1. Activate your virtual environment with "venv\Scripts\activate" if not already activated.
   2. type in "python main.py" and hit enter to run the program on all spreadsheets in the "Excel Files folder"  
   Alternatively, you can use "python main.py -f "filename" to run the program on a single file.  
   Note: This file would be in your project folder, not the Excel Files subfolder. 
   3. Once you are done executing the program, type "deactivate" to close your virtual environment