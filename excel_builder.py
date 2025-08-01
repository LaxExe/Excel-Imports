from openpyxl import Workbook
import json 
from row_parsing import gather_row_data

def export_to_excel(data, output_file):
    wb = Workbook()
    ws = wb.active

    if data:
        # Write headers from the keys of the first dictionary
        headers = list(data[0].keys())
        ws.append(headers)

        # Write the row values
        for row_dict in data:
            ws.append([row_dict.get(header, "") for header in headers])
    else:
        # Add placeholder header so the file isn't completely empty
        ws.append(["  "])

    wb.save(output_file)
    print(f"Excel file saved as: {output_file}")
