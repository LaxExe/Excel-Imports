# Takes in an Excel file
# Reads the first 5 rows of data, and runs a loop to find last colum needed
# Create a new Excel file with the data

# Feed Data into OpenAI API
# via prompt we will ask it for informatin about the data
    # Identify the required columns
        # Phone Number
        # Email
        # Full Name
        # Adress
            # Adress Format
                # Adress Information
        # Identify Feilds with additional Information

# Generate Json Object with the following feilds

from openpyxl import load_workbook, Workbook
from openai import OpenAI
from dotenv import load_dotenv
import os
import json



load_dotenv()
api_key = os.getenv("OPENAI_API_KEY")

client = OpenAI(
    api_key=api_key,
    base_url="https://api.sambanova.ai/v1"
)


def get_first_5_rows_as_list(file_path):
    wb = load_workbook(file_path)
    ws = wb.active
    rows = []

    row_index = 1
    for row in ws.iter_rows(values_only=True):
        if row_index > 5:
            break

        row_data = []
        for cell in row:
            row_data.append(str(cell) if cell is not None else "")
        
        rows.append(row_data)
        row_index += 1

    data_string = json.dumps(rows)

    return data_string


primary_excel_prompt = """You will be give an Excel file information with the first 5 rows of data, 
your job is to fill out all the feilds in the JSON object below YOU MUST FILL OUT ALL OF THEM.
There should be no formating or any other explanation, just the JSON object. The JSON object should be in the following format:

"""

primary_json_structure = """
  "required_fields": {
    "phone_number": "Column Number",
    "email": "Column Number",
    "full_name": "Column Number",
    "address_format": "1 Column | multi-Column"
  },
  "additional_fields": {
    "Column Name 1": "Column Number",
    "Column Name 2": "Column Number"
  }
}
"""

def AI_generate_json_structure(data):


    response = client.chat.completions.create(
                model="Meta-Llama-3.1-8B-Instruct",
                messages=[
                    {
                        "role": "user",
                        "content": primary_excel_prompt + primary_json_structure + data
                    }
                ]
            )
    
    response_content = response.choices[0].message.content.strip()

    try:
        info_json = json.loads(response_content)
        required_fields = ["phone_number", "email", "full_name", "address_format"]
        
        for field in required_fields:
            value = info_json.get(field)
            if value in [None, "", "null"]:
                print(f"Missing required field '{field}' in the excel file.")
                all_required_present = False
                break

        
        if all_required_present:
            with open(f"info.json", "w") as f:
                json.dump(info_json, f, indent=4)
                
    except json.JSONDecodeError:
        pass

    return info_json



single_column_address_prompt = """
you will be give excel file data with the first 5 rows of data, in this data the address
is arranged in a single column you job is to fill out the JSON object below with the correct information about the address
you must fill out all the fields in the JSON object below with the way that the address are arranged
you should alos identify the seperator and corectly put in the format, there should be no formating or any other explanation, just the JSON object.
"""

single_column_address_format = """

{
  "address_type": "1 Column",
  "address_format": "country, city, postal, street, unit",
  "address_separator": ","
}

"""

multicolumn_address_prompt = """ 
you will be give excel file data with the first 5 rows of data, in this data the address
is arranged in multiple columns you job is to fill out the JSON object below with the correct column numbers
you must fill out all the fields in the JSON object below, there should be no formating or any other explanation, just the JSON object.
Make sure to use the correct column numbers for each field, if there are no columns with that feild, you must put a null value.
"""

multicolumn_address_format = """
{
  "address_type": "multi-Column",
  "address": {
    "street_address": "Column Number",
    "city": "Column Number",
    "postal": "Column Number"
  },
  "additional_address_information": {
    "province": "Column Number",
    "country": "Column Number"
  }
}

"""

def identify_address_format(data, info_json):
    if(info_json.get("address_format") == "multi-Column"):

        response = client.chat.completions.create(
                model="Meta-Llama-3.1-8B-Instruct",
                messages=[
                    {
                        "role": "user",
                        "content": multicolumn_address_prompt + multicolumn_address_format + data
                    }
                ]
            )
        
        response_content = response.choices[0].message.content.strip()

        try:
            address_json = json.loads(response_content)
            required_fields = ["street_address", "city", "postal"]
            
            for field in required_fields:
                value = address_json.get(field)
                if value in [None, "", "null"]:
                    print(f"Missing required field '{field}' in the multi address identification.")
                    all_required_present = False
                    break

            
            if all_required_present:
                with open(f"Multiaddress.json", "w") as f:
                    json.dump(info_json, f, indent=4)
                    
        except json.JSONDecodeError:
            pass

    else:
        response = client.chat.completions.create(
                model="Meta-Llama-3.1-8B-Instruct",
                messages=[
                    {
                        "role": "user",
                        "content": single_column_address_prompt + single_column_address_format + data
                    }
                ]
            )
        
        response_content = response.choices[0].message.content.strip()


        try:
            address_json = json.loads(response_content)
            required_fields = ["address_type", "address_format", "address_separator"]
            
            for field in required_fields:
                value = address_json.get(field)
                if value in [None, "", "null"]:
                    print(f"Missing required field '{field}' in the single address identification.")
                    all_required_present = False
                    break

            
            if all_required_present:
                with open(f"Multiaddress.json", "w") as f:
                    json.dump(info_json, f, indent=4)
                    
        except json.JSONDecodeError:
            pass


    return address_json
        

