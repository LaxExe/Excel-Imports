from openpyxl import Workbook
import json 
from row_parsing import gather_row_data

def export_to_excel(data, output_file):
    if not data:
        print("No data to write.")
        return

    wb = Workbook()
    ws = wb.active

    # Write headers from the keys of the first dictionary
    headers = list(data[0].keys())
    ws.append(headers)

    # Write the row values
    for row_dict in data:
        ws.append([row_dict.get(header, "") for header in headers])

    wb.save(output_file)
    print(f"Excel file saved as: {output_file}")


with open("info.json", "r") as f:
  info_json = json.load(f)


