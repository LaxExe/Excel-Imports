from Excel_interpreter import get_first_5_rows_as_dict, AI_generate_json_structure
import json
from row_parsing import gather_row_data
from Excel_builder import export_to_excel


snipit = get_first_5_rows_as_dict("test.xlsx")
print(snipit)
AI_generate_json_structure(snipit)


with open("info.json", "r") as f:
    info_json = json.load(f)

results = gather_row_data('test.xlsx', info_json)
export_to_excel(results, "output.xlsx")