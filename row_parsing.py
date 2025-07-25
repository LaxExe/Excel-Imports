from openpyxl import load_workbook
import json
from clean import validate_email, filter_full_name, clean_phone_number, extract_country_code

# Helper Function
def col_letter_to_index(letter):
  if not letter or letter == None:
    return None
  total = 0
  for i, c in enumerate(reversed(letter.upper())):
    total += (ord(c) - 64) * (26 ** i)
  return total




def gather_row_data(excel_file, json_structure):
  
  wb = load_workbook(excel_file)
  ws = wb.active

  results = []



  col_indexes = {
    "email": col_letter_to_index(json_structure["required_fields"]["email"]),
    "phone_number": col_letter_to_index(json_structure["required_fields"]["phone_number"]),
    "full_name": col_letter_to_index(json_structure["required_fields"]["full_name"]),
    "last_name": col_letter_to_index(json_structure["required_fields"]["last_name"])
  }

  for i in range(2, ws.max_row + 1):

    email = ws.cell(row=i, column=col_indexes["email"]).value
    phone_number = ws.cell(row=i, column=col_indexes["phone_number"]).value
    full_name = ws.cell(row=i, column=col_indexes["full_name"]).value

    # Case 1: Last Name Provided but None
    if (col_indexes["last_name"] == None):
      validate_name = filter_full_name(full_name, email, col_indexes["last_name"])  
    else:
      last_name = ws.cell(row=i, column=col_indexes["last_name"]).value
      validate_name = filter_full_name(full_name, email, last_name)

    
    valid_email = validate_email(email)
    valid_phone_number = clean_phone_number(phone_number, validate_name, valid_email, None) # Enter the country at none 

    results.append({
      "email": valid_email,
      "phone_number": valid_phone_number,
      "full_name": validate_name
    })

  
  return results




  
 
with open("info.json", "r") as f:
    info_json = json.load(f)

results = gather_row_data('test.xlsx', info_json)
print(results)