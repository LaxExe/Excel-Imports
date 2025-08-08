from openpyxl import Workbook
import json 

def safe_str(val):
    if isinstance(val, (dict, list)):
        return json.dumps(val)
    # Handle None values
    if val is None:
        return ""
    return str(val)

def export_to_excel(data, output_file):
    wb = Workbook()
    ws = wb.active

    if data:
        # Write headers from the keys of the first dictionary
        headers = list(data[0].keys())
        ws.append(headers)

        # Write the row values
        for row_dict in data:
            ws.append([safe_str(row_dict.get(header, "")) for header in headers])
    else:
        headers = ["email", "phone_number", "full_name", "shipping_address", "billing_address", "additional_fields"]
        ws.append(headers)

    wb.save(output_file)
    print(f"Excel file saved as: {output_file}")
