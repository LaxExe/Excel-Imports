
from openpyxl import load_workbook
import json

# Helper Function
def col_letter_to_index(letter):
  if not letter:
    return None
  total = 0
  for i, c in enumerate(reversed(letter.upper())):
    total += (ord(c) - 64) * (26 ** i)
  return total



# Validation Function
def filter_full_name(name, email, lastname):
  
  # Case 2
  if name and lastname:
    return f"{name.strip().title()} {lastname.strip().title()}"
 
  # Case 1
  if name and " " in name: # multi part name 
        name_parts = name.strip().split()
        return " ".join(part.capitalize() for part in name_parts)

  # Case  3
  if name == None and lastname == None and email:
    name_segments = email.split("@")[0]
    words = name_segments.replace(".", "").replace("_", " ").split()
    return " ".join(word.capitalize() for word in words)
  
  return None


def gather_row_data(excel_file, json_structure):
  
  wb = load_workbook(excel_file)
  ws = wb.active

  results = []

  col_indexes = {
    "email": col_letter_to_index(json_structure["required_fields"]["email"]),
    "phone_number": col_letter_to_index(json_structure["required_fields"]["phone_number"]),
    "full_name": col_letter_to_index(json_structure["required_fields"]["full_name"]),
  }

  for i in range(2, ws.max_row + 1):

    email = ws.cell(row=i, column=col_indexes["email"]).value
    phone_number = ws.cell(row=i, column=col_indexes["phone_number"]).value

    full_name = ws.cell(row=i, column=col_indexes["full_name"]).value
    validate_name = filter_full_name("Bob", email, "tim")
    
    results.append({
      "email": email,
      "phone_number": phone_number,
      "full_name": validate_name
    })

  
  return results



with open("info.json", "r") as f:
  info_json = json.load(f)

results = gather_row_data('test.xlsx', info_json)
print(results)

  
 
