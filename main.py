from Excel_interpreter import get_first_5_rows_as_dict, AI_generate_json_structure

snipit = get_first_5_rows_as_dict("canadian_client_data.xlsx")
print(snipit)
AI_generate_json_structure(snipit)

# Work on this later:
"""
results = gather_row_data('test.xlsx', info_json)
export_to_excel(results, "output_results.xlsx")
"""
