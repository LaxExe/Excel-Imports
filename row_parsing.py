from openpyxl import load_workbook
import json
from clean import validate_email, filter_full_name, clean_phone_number, AI_check
from address import is_1_column_tag, column_1_address_skip, column_multi_address

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
  check = 0
  page = 0
  
  wb = load_workbook(excel_file)
  ws = wb.active
  good_data = []

  results = []
  failed_results = []

  required_format = json_structure["address_1_column_format"]["ideal_address_format"]

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


      
      valid_address = column_1_address_skip(address, address_1_column_format, address_1_column_seperator, required_format)       


      if (valid_address == True or valid_email == False or valid_phone_number == False  or validate_name == False):
        failed_results.append (   
        f'{{\n  "email": "{email}",\n  "phone_number": {valid_phone_number},\n  "full_name": "{full_name}",\n  "address": "{address}",\n  "additional_fields": "{valid_items}"\n}}')
        check = check + 1



      if valid_address != True:

        if sample < 3:
          good_data.append (   
          f'{{\n  "email": "{email}",\n  "phone_number": {valid_phone_number},\n  "full_name": "{full_name}",\n  "address": "{address}",\n  "additional_fields": "{valid_items}"\n}}'
          )
          sample = sample + 1


        results.append({
          "email": valid_email,
          "phone_number": valid_phone_number,
          "full_name": validate_name,
          "address": valid_address,
          "additional_feilds" : valid_items
        })
  
      if check == 10:
        AI_check(good_data, failed_results, page)
        failed_results = []
        check = 0 
        page = page + 1
  
  else:
    # CASE 2: 
    # This handles Excel files where address information is spread across multiple columns
    # Instead of one "Address" column, we have separate columns for street, city, province, etc.
    
      for i in range(2, ws.max_row + 1):
        # Create a dictionary to hold all the cell values for this row
        # Format: {"A2": "john@email.com", "B2": "123 Main St", "C2": "Toronto", ...}
        row_dict = {}
        
        # Iterate through each column in this specific row
        # This builds our row_dict with all the cell values for processing
        for col in ws.iter_cols(min_row=i, max_row=i):
          cell = col[0]  # Get the cell from this column at row i
          # Store the cell value with its reference (e.g., "A2", "B2") as the key
          # Convert to string and strip whitespace, default to empty string if cell is None
          row_dict[f"{cell.column_letter}{cell.row}"] = str(cell.value).strip() if cell.value else ""

        # Extract the required fields using the JSON structure mapping
        # The JSON tells us which columns contain email, phone, name, etc.
        email = row_dict.get(f"{json_structure['required_fields']['email']}{i}", "")
        phone_number = row_dict.get(f"{json_structure['required_fields']['phone_number']}{i}", "")
        full_name = row_dict.get(f"{json_structure['required_fields']['full_name']}{i}", "")
        
        # Handle last name field - it might not exist in the JSON structure
        # If last_name is null in JSON, we pass None to the name validation function
        last_name_val = (
          row_dict.get(f"{json_structure['required_fields']['last_name']}{i}", "")
          if json_structure["required_fields"]["last_name"]
          else None
        )
        
        # Validate and format the full name using the name cleaning function
        # This handles cases where we have separate first/last name columns
        validate_name = filter_full_name(full_name, email, last_name_val)

        # Validate the email address using the email cleaning function
        # This ensures the email is properly formatted and valid
        valid_email = validate_email(email)

        # Clean and format the address using the multi-column address function
        # This is the key step for multi-column addresses - it combines separate columns into one address
        address_result = column_multi_address(row_dict, i, json_path="info.json")
        formatted_address = address_result["address"]  # The complete formatted address string
        needs_ai = address_result["needs_ai"]          # Which address parts need AI to fill in

        # Use the formatted address for phone number validation
        # If the address is complete (no AI needed), we can use it to help validate the phone number
        # If address needs AI, we pass None so phone validation doesn't rely on incomplete address
        address_for_phone = formatted_address if not needs_ai else None
        valid_phone_number = clean_phone_number(phone_number, validate_name, valid_email, address_for_phone)

        # Process additional fields (any columns that aren't email, phone, name, or address)
        # These might be things like "Unit ID", "Owner", "Department", etc.
        items = []
        additional_fields = json_structure.get("additional_fields", {})

        # Loop through each additional field and get its value from the row
        for key, col_letter in additional_fields.items():
          cell_key = f"{col_letter}{i}"  
          val = row_dict.get(cell_key, "")  
          items.append(f"{key}: {val}")  
        
        # Join all additional fields with triple spaces for readability
        # This creates a string like "Unit ID: asdf123   Owner: Nice"
        valid_items = '   '.join(items)

        # Determine if this row needs to be sent to AI for processing
        # A row needs AI if any of these conditions are true:
        # - Email is invalid
        # - Phone number is invalid  
        # - Name is invalid
        # - Address needs AI (missing or invalid address parts)


        results.append({
          "email": valid_email,
          "phone_number": valid_phone_number,
          "full_name": validate_name,
          "address": formatted_address,
          "additional_feilds" : valid_items
        })
  

    # Send the good examples and failed results to AI for processing
    # The AI will try to fix any issues in the failed results using the good examples as reference
  AI_check(good_data, failed_results, page)
  return results




