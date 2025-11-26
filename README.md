# Abbott-Excel

This program uses python to clean up Abbott spreadsheet data.  
It can clean up, reorganize, and perform regression analysis on all .xlsx files inside the "Excel Files" subfolder and export them to a newly created "Finished Files" subfolder.
The original files will not be modified. 

## For your first time:
Please download python onto your computer before using (https://www.python.org/downloads/)
1. Click "Download Python 3.x.x", using the latest stable version
2. Run the installer, remember to check "Add Python.exe to PATH" at the bottom, then "Install Now"
3. Verify python installation. Open command prompt and type "python --version"

## Creating a Virtual Environment and Installing Dependencies
You will also need to create a virtual environment and install some python modules. This will prevent the program from affecting anything else on your computer:
1. Create and navigate to the folder that you will use for this program.  
Right click the folder in file explorer, and click "copy as path"  
Open the command prompt from your windows start menu, type "cd " and paste in the file path. Then hit enter
2. Enter in "python -m venv venv" to create your virtual environment
3. Enter in "venv\Scripts\activate" to activate your virtual environment. You should see a (venv) text in your command line after.
4. Enter in "python -m pip install --upgrade pip"
5. Enter in "pip install -r requirements.txt"
7. Enter in "deactivate" to deactivate the virtual environment


## To use the program:
1. Download the requirements.txt, main.py, cleaning.py, analysis.py, and config_headers.py files from the github repository into your project folder
2. Create a subfolder called "Excel Files" within your project folder and store any spreadsheets you want to edit inside
3. Copy into the project folder the files that will be utilized for cleaning (ex. formula codes, nutrient list, etc.)  
   Note: these same helper files will be used for all data files within the "Excel Files" subfolder
3. Modify the config_headers.py file with a text editor to match your file, tab, and column header names. Note that the program will use the same formatting for all spreadsheets. 
4. Open up the command line from the windows start menu, navigate to the project folder following the instruction 1 in the virtual environment section above
   1. Activate your virtual environment with "venv\Scripts\activate" if not already activated.
   2. type in "python main.py" and hit enter to run the program on all spreadsheets in the "Excel Files folder"  
   Note: by default, using "python main.py" will only clean and re-organize data.  
   Alternatively you can use:
      1. "python main.py -b" - this will run **both cleaning and regression analysis** on the files within "Excel Files"
      2. "python main.py -a" - this will run **only regression analysis** on the files within "Excel Files"
   3. Once you are done executing the program, type "deactivate" to close your virtual environment