import json
import os
import re
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from address import street_and_postal_code, street_and_city, fill_missing_address_with_geopy
 
def find_last_row_with_data(ws):
  for row in reversed(range(2, ws.max_row + 1)):
    if any(cell.value not in (None, "") for cell in ws[row]):
      return row
  return 1 # First row if no data found
 
def extract_remake_number(filename):
    # Extract number from remake#.json to sort numerically
    match = re.search(r'remake(\d+)\.json', filename)
    return int(match.group(1)) if match else float('inf')
 
def append_cleaned_json_to_excel(directory, output_excel):
 
  results = []
  # If output_excel doesn't exist, exit the function



 
  # Otherwise, load the existing workbook
  wb = load_workbook(output_excel)
  ws = wb.active
 
  # grab headers from the first row
  headers = [cell.value for cell in ws[1]] # travels linearly through the first row
 
  # Row thingy
  last_data_row = find_last_row_with_data(ws)
  next_row = last_data_row + 1
 
  # Seperator shade - apply shade periodically via for loop and update the next_row variable
  gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
  for col in range(1, len(headers) + 1):
    ws.cell(row=next_row, column=col).fill = gray_fill
  next_row += 1
 
  # Get and sort remake files numerically
  remake_files = [
    f for f in os.listdir(directory)
    if re.match(r'remake\d+\.json$', f)
  ]
  print("Remake files found:", remake_files)
  sorted_files = sorted(remake_files, key=extract_remake_number)
 
  # Append data from each JSON file in directory
  for file in sorted_files:
    with open(os.path.join(directory, file), "r") as f:
      data = json.load(f)
 
    # Get clean items list
    clean_items = data.get("clean_items", []) # -> Fall back to empty list if not found
    for item in clean_items:
 
      address_info = item.get("address")[0]
      street = address_info.get("street_address", None)
      city = address_info.get("city", None)
      postal = address_info.get("postal_code", None)
 
      if not item.get("missing_parts_of_address", True):
        address_info = item.get("address")[0]
        full_address = f"{street}, {postal}, {city}, {address_info.get('province_or_state_name')}, {address_info.get('country', "")}"

      # TODO Fix this as it returns none, when the feild is null
     
      # Else we're dealing with missing items in Address feild
      else:
        
        if city == None and street and postal:
          full_address = street_and_postal_code(street, postal)
        elif postal == None and street and city:
          full_address = street_and_city(street, city)
        elif street and city and street:
          full_address = street_and_postal_code(street, postal)
          if full_address == None:
            full_address = street_and_city(street, city)
        else:
          full_address = "Missing"
 
 
      row_data = {
        "email": item.get("email", "value-missing"),
        "phone_number": item.get("phone_number", "value-missing"),
        "full_name": item.get("full_name", "value-missing"),
        "address": full_address,
        "additional_feilds" : item.get("additional_fields", "value-missing")
      }
 
      ws.append([row_data.get(header, "") for header in headers])
      results.append(row_data)
      next_row += 1
 
  wb.save(output_excel)
  print(results)
  print(f"Data appended to {output_excel} successfully from {len(sorted_files)} remake files.")
 
