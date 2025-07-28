from openpyxl import load_workbook
import json
from clean import validate_email, filter_full_name, clean_phone_number, AI_check
from address import is_1_column_tag, column_1_address_skip

# Helper Function
def col_letter_to_index(letter):
  if not letter or letter == None:
    return None
  total = 0
  for i, c in enumerate(reversed(letter.upper())):
    total += (ord(c) - 64) * (26 ** i)
  return total




def gather_row_data(excel_file, json_structure):

  sample = 0
  
  wb = load_workbook(excel_file)
  ws = wb.active
  good_data = []

  results = []
  failed_results = []

  if is_1_column_tag("info.json"):

    additional_feilds = json_structure["additional_fields"]

    col_indexes = {
      "email": col_letter_to_index(json_structure["required_fields"]["email"]),
      "phone_number": col_letter_to_index(json_structure["required_fields"]["phone_number"]),
      "full_name": col_letter_to_index(json_structure["required_fields"]["full_name"]),
      "last_name": col_letter_to_index(json_structure["required_fields"]["last_name"]),
      "address" : col_letter_to_index(json_structure["address_1_column_format"]["address_column"])
    }

    additional_fields = json_structure.get("additional_fields", {})

    for key, col_letter in additional_fields.items():
        col_indexes[key] = col_letter_to_index(col_letter)

    address_1_column_format = json_structure["address_1_column_format"]["address_format"]
    address_1_column_seperator = json_structure["address_1_column_format"]["address_separator"]

    for i in range(2, ws.max_row + 1):
      
      address = ws.cell(row=i, column=col_indexes["address"]).value
      email = ws.cell(row=i, column=col_indexes["email"]).value


      # Get phone number from excel
      phone_number = ws.cell(row=i, column=col_indexes["phone_number"]).value
      
      
      full_name = ws.cell(row=i, column=col_indexes["full_name"]).value

      # Case 1: Last Name Provided but None
      if (col_indexes["last_name"] == None):
        validate_name = filter_full_name(full_name, email, col_indexes["last_name"])  
      else:
        last_name = ws.cell(row=i, column=col_indexes["last_name"]).value
        validate_name = filter_full_name(full_name, email, last_name)

      items = []

      for key in additional_feilds:
        items.append(key + " : "+ (ws.cell(row=i, column=col_indexes[key])).value + "   ")

      valid_items = ' '.join(items)

      valid_email = validate_email(email)

      # call on the valid phone number method
      valid_phone_number = clean_phone_number(phone_number, validate_name, valid_email, None) # Enter the country at none 



      valid_address = column_1_address_skip(address, address_1_column_format, address_1_column_seperator)       


      if (valid_address == True or valid_email == False or valid_phone_number == False  or validate_name == False):
        failed_results.append (   
          f"{email}, 416 715 6897, {full_name}, {address}, {valid_items}" # replace with  {valid_phone_number}
        )



      if valid_address != True:

        if sample < 5:
          good_data.append (   
          f"{valid_email},  416 715 6897, {validate_name}, {valid_address}, {valid_items}"
        )
          sample = sample + 1


        results.append({
          "email": valid_email,
          "phone_number": valid_phone_number,
          "full_name": validate_name,
          "address": valid_address,
          "additional_feilds" : valid_items
        })
  
  AI_check(good_data, failed_results)
  

  return results




  
 
with open("info.json", "r") as f:
    info_json = json.load(f)

results = gather_row_data('test.xlsx', info_json)
print(results)