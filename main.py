from Excel_interpreter import get_first_5_rows_as_dict, AI_generate_json_structure
import json
from row_parsing import gather_row_data
from excel_builder import export_to_excel
from jsons_to_excel import append_cleaned_json_to_excel


# _________________________________________________________________________________________________
#
# ----------------------------------------- WORKFLOW ----------------------------------------------
# _________________________________________________________________________________________________



# 1. Extract First 5 Rows for AI to understand the columns and format
snipit = get_first_5_rows_as_dict("test.xlsx")
print(snipit)

# 2. Send the snipit to Ai - creates info.json - maps each category to its respective Column Letter
AI_generate_json_structure(snipit)

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
