from Excel_interpreter import get_first_5_rows_as_dict, AI_generate_json_structure
import json
from row_parsing import gather_row_data
from excel_builder import export_to_excel
from jsons_to_excel import append_cleaned_json_to_excel


# _________________________________________________________________________________________________
#
# ----------------------------------------- WORKFLOW 2 --------------------------------------------
# _________________________________________________________________________________________________


# 3. If the cell data are validated and formatted without returning null -> Process them in a list of dictionaries 
    # 3b. Store any failed results and send to AI -> Ai makes remake.json
with open("info.json", "r") as f:
    info_json = json.load(f)

results = gather_row_data('test.xlsx', info_json)
print("Results:", results)

# 4. Validated Data is exported to the excel
export_to_excel(results, "output.xlsx")

# 5. Iterate through remake.json, fix addresses with Geopy, append to existing output excel
append_cleaned_json_to_excel(directory=".", output_excel="output.xlsx")


# ISSUES:
# 1. One row of information merges several rows, returns None and Missing for all properties
# 2. Names are returning None when email not present
# 3. Phone Number is either none or empty string
# 4. Emails are Missing when not provided
# 5. Address can either be missing or unable to extract address
# 6. Additional Info is None

# Task:

# What is the core item that makes it valid to append to excel? --> Email & Phone Number
# Store data that is missing into a dictionary where the key is row number, the value is missing feild (conditional on if email or phone number is present)
# Highlight Empty "MISSING" data in Red. 

# RESULT:
# Two ways to view missing data: Excel and Dictionary
# know what to append and what not to append bc rn everything is being appended, even empty rows