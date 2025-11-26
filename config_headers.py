# --------- MODIFY THESE FOLDER/FILE/TAB NAMES AS NEEDED ----------
# ***** All file/sheet/header names are CASE + SPACE sensitive *****

# ----- Folder Names -----

# folder that has the original data files
data_folder = "Excel Files"

# folder name where modified files will be stored
finshed_folder = "Finished Files"

# folder name where helper files (formula code, nutrients, etc.) are stored
helper_folder = "Helper Files"

# ----- End of Folder Names -----


# ------- Excel Tab/Sheet Names -------

data_s = 'Data'     # this is the tab where the data is stored in each Excel file


# ----- Start of Unit Conversion Section -----

# --- Primary unit conversions ---

# this is the file and tab name for primary unit conversion
first_conv_file = 'Unit conversion - 2.xlsx'
first_conv_s = 'Sheet1'

# this is the # of rows in the primary conversion to skip before headers
first_conv_skip = 2

# these are the headers inside the primary unit conversion file
first_conv_test_h = "Test"
first_conv_units_h = 'Unit'
first_conv_h = "Conversion factor"

# --- end of primary unit conversion ---

# secondary unit conversions
second_conv_file = '1 - Data Cleansing.xlsx'
unit_s = '4-UnitConversion'

# ----- End of Unit Conversion section -----


# --- Formula Code section ---

# this is the file and tab that contains the formula codes
formula_code_file = '1 - Data Cleansing.xlsx'
form_s = '3-Formulacode'

# invalid formula codes (don't need the whole name, just a part is fine)
invalid_formulas = ["Can't find", "Do not use", "Pending", "Ask", "Too new", "development"]

# --- End of Formula Code section ---


# --- Nutrient List section ---

# this is the file and tab that contains the nutrients list
nutrient_file = '6 - Nutrient List.xlsx'
nutrient_s = 'Nutrient List'

# how many rows in nutrient list file to skip before headers
nutrient_skip = 1

# which columns in nutrient list to use, Col A = 0, B = 1...
nutrient_cols = [0, 1, 3]

# --- End Nutrient List section ---


# this is the file and tab for the VitE Factor
vitE_file = '1 - Data Cleansing.xlsx'
vitE_s = '5-VitEFactor'


# ------- end of Excel tab/sheet names -------


# --- Header Names from the Data + extra tabs ---

batch_h = 'BATCH'
project_h = 'PROJECT'
run_h = 'RUN'
prod_date_h = 'PRODUCTION_DATE'
comp_date_h = 'DATE_COMPLETED'
storage_h = 'STORAGE_CONDITION'
dur_h = 'AB_INTERVAL'
analysis_h = 'ANALYSIS'
name_h = 'NAME'
unit_h = 'UNITS'
conversion_formula_h = 'Formula Code '
conversion_sources_h = 'Sources'
conversion_test_h = "Test"
conversion_units_h = 'Units'
conversion_conv_h = "Conversion factor"
header_h = 'Heading for the new set of columns'
ab_stage_h = 'AB_STAGE'
description_h = 'DESCRIPTION'
batch_type_h = 'BATCH_TYPE'
batch_category_h = 'BATCH_CATEGORY'
manu_loc_h = 'MANUFACTURING_LOCATION'
ab_container_h = 'AB_CONTAINER'
text_h = 'TEXT'

# --- End of Data tab header names ---


# -------- New File/Sheet/Header Names --------

# prefix to modified Excel File name (this can be empty - '')
prefix = 'updated_'


# --- Newly Created Sheet Names ---

cleaned_s = 'CleanedData'       # sheet name for cleaned data
organized_s = 'Re-Organized Data'   # sheet name for re-organized (test as headers) data
regression_s = 'Regressions'    # sheet name for regression analysis
average0_s = '0 Time Average'
average12_s = '12 Month Average'

# --- header names in new sheets ---
temp_h = "Temperature"
humidity_h = "Humidity"
interval_h = "Interval (D)"
form_h = 'Formula Code'
sources_h = 'Sources'
test_h = 'Test'
newUnit_h = 'Units'
conv_h = 'Conversion Factor'
vitE_h = 'Vit E factor'
results_h = 'Results'

# --- data type header/value ---
data_type_h = 'Data Type'
data_type_value = 'LIMS Test'

t0_h = "Result at t=0M"
t12_h = "Result at t=12M"
percent12_h = "% Remaining at t=12M"
avg12_h = "Avg % at t=12M"
avg_h = 'Average'   # header for average calculation
