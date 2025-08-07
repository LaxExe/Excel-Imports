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

def build_full_address(address_info):
    if not address_info:
        return "Missing"

    street = address_info.get("street_address", "")
    city = address_info.get("city", "")
    postal = address_info.get("postal_code", "")
    province = address_info.get("province_or_state_name", "")
    country = address_info.get("country", "")

    if street and city and postal:
        return f"{street}, {city}, {province}, {country}, {postal}"

    # Geopy fallback logic
    if not city and street and postal:
        return street_and_postal_code(street, postal)
    elif not postal and street and city:
        return street_and_city(street, city)
    elif street and (city or postal):
        return street_and_postal_code(street, postal or "")
    
    return "Missing"



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
    print(f"Processing {file}...")

    count = 0
    # Get clean items list
    clean_items = data.get("clean_items", []) # -> Fall back to empty list if not found
    for item in clean_items:
      print("Processing row:", count) 
      count += 1
      shipping_address_info = (item.get("shipping_address") or [{}])[0]
      billing_address_info = (item.get("billing_address") or [{}])[0]

      if shipping_address_info != billing_address_info: 
        
        full_shipping_address = build_full_address(shipping_address_info)
        full_billing_address = build_full_address(billing_address_info)
      
      else:
        full_shipping_address = shipping_address_info.get("full_address") or build_full_address(shipping_address_info)
        full_billing_address = billing_address_info.get("full_address") or build_full_address(billing_address_info)


      row_data = {
        "email": item.get("email", "value-missing"),
        "phone_number": item.get("phone_number", "value-missing"),
        "full_name": item.get("full_name", "value-missing"),
        "shipping_address": full_shipping_address,
        "billing_address": full_billing_address,
        "additional_fields" : item.get("additional_fields", "value-missing")
      }
 
      ws.append([row_data.get(header, "") for header in headers])
      results.append(row_data)
      next_row += 1
 
  wb.save(output_excel)
  print(results)
  print(f"Data appended to {output_excel} successfully from {len(sorted_files)} remake files.")
 
