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
