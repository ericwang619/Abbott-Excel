# --------- MODIFY THESE TAB/HEADER NAMES AS NEEDED ----------

# Excel Tab/Sheet Names
data_s = 'Data'     # this is the tab where the data is stored in each excel file

# this is the file/tab that contains the formula codes
formula_code_file = 'instructions.xlsx'
form_s = '3-Formulacode'

# this is the file/tab for the VitE Factor
vitE_file = 'instructions.xlsx'
vitE_s = '5-VitEFactor'

# this file/tab contains the nutrients list
nutrient_file = 'instructions.xlsx'
nutrient_s = '6-NutrientList'

# this is the file/tab that contains the primary unit conversions
first_conv_file = 'Unit conversion - 2.xlsx'
first_conv_s = 'Sheet1'
first_conv_test_h = "Test"
first_conv_units_h = 'Unit'
first_conv_h = "Conversion factor"

# this is the file/tab used for unit conversion after checking the primary
second_conv_file = 'instructions.xlsx'
unit_s = '4-UnitConversion'



# Header Names from the Data + extra tabs
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
ab_stage_h = 'AB_STAGE'
description_h = 'DESCRIPTION'
batch_type_h = 'BATCH_TYPE'
batch_sub_h = 'BATCH_SUB_CATEGORY'
manu_loc_h = 'MANUFACTURING_LOCATION'
ab_container_h = 'AB_CONTAINER'
text_h = 'TEXT'


# invalid formula codes
invalid_formulas = ["Can't find", "Do not use", "Pending", "Ask", "Too new", "development"]


# ------------- DO NOT MODIFY BELOW THIS LINE ------------

# newly created sheet names
updated_s = 'UpdatedData'
consolidated_s = 'ConsolidatedData'
stats_s = 'FormulaStats'
clusters_s = 'Clusters'

# new headers
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

new_headers = [temp_h, humidity_h, interval_h, form_h, sources_h,
               test_h, newUnit_h, conv_h, vitE_h, results_h]

data_type_h = 'Data Type'


avg_h = 'Average'
min_h = 'Minimum'
max_h = 'Maximum'
count_h = 'Count'



data_sheet_name = "1 - Data Cleansing.xlsx"