import re
import phonenumbers
import pycountry
from dotenv import load_dotenv
import os
from openai import OpenAI
import json

with open("info.json", "r") as f:
    data = json.load(f)         
    json_string = json.dumps(data)  




load_dotenv()
api_key = os.getenv("OPENAI_API_KEY")

client = OpenAI(
    api_key=api_key,
    base_url="https://api.sambanova.ai/v1"
)




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


def validate_email(email):
  if not email or "@" not in email:
    return None
  return email.strip().lower()



def extract_country_code(address):
    """
    Extract the country name from the address and convert to region code.
    """
    if not address:
        return None

    for country in pycountry.countries:
        if country.name.lower() in address.lower():
            return country.alpha_2  # e.g., 'US', 'IN', etc.
    return None


def clean_phone_number(raw_phone, full_name, email, address):
    """
    Cleans and standardizes a phone number:
    - Keeps the country code if present
    - Infers country code from address if not present
    - If phone is empty but name and email are present, returns empty string
    - Returns phone number in +E.164 format if valid
    """

    # If phone is empty and name/email present - return ""
    if (not raw_phone or str(raw_phone).strip() == "") and full_name.strip() and email.strip():
        return ""

    phone_str = str(raw_phone).strip()
    # Remove all non-numeric characters except '+'
    phone_str = re.sub(r"[^\d+]", "", phone_str)

    # If phone starts with '+', parse it directly
    if phone_str.startswith("+"):
        try:
            parsed = phonenumbers.parse(phone_str, None)
            if phonenumbers.is_valid_number(parsed):
                return phonenumbers.format_number(parsed, phonenumbers.PhoneNumberFormat.E164)
        except phonenumbers.NumberParseException:
            return ""
    else:
        # Try to infer region from address
        region = extract_country_code(address) or "US"
        try:
            parsed = phonenumbers.parse(phone_str, region)
            if phonenumbers.is_valid_number(parsed):
                return phonenumbers.format_number(parsed, phonenumbers.PhoneNumberFormat.E164)
        except phonenumbers.NumberParseException:
            return ""

    return ""

def clean_ai_response(response_text: str) -> str:
    text = response_text.strip()
    if text.startswith("```"):
        first_newline = text.find("\n")
        if first_newline != -1:
            text = text[first_newline + 1:]
        if text.endswith("```"):
            text = text[:-3]
    return text.strip()


clean_prompt = """
You will be given two lists of strings containing client information entries.

The first list contains **correctly formatted entries**. 
The second list contains **entries with formatting issues** such as:
- missing or extra commas,
- extra or inconsistent spaces,
- incorrect word ordering,
- malformed or abbreviated address parts,
- or other structure errors.

You will also be given a JSON structure describing the desired field order.

Your task is to correct the formatting of the second list ("bad" entries) to match the **exact style and structure** of the "good" list. Try to infer and apply the same:
- field order,
- punctuation (especially commas),
- capitalization,
- spacing.
- Your goal is to make the bad entries to be the same as a the good entries

Be precise and conservative. **Do not invent or replace values**—only clean and rearrange as needed to make them structurally identical to the examples in the good list.

Return **only** a corrected dictionary values in the exact format used in the good list, and nothing else—no extra text, no explanation, no formatting.

        {
          "email": valid_email,
          "phone_number": valid_phone_number,
          "full_name": validate_name,
          "address": valid_address,
          "additional_feilds" : valid_items
        }

The list of correct examples is:
"""

def AI_check(results, failed_results):
    
  string_good_data = str(results)
  string_failed_results = str(failed_results)

  print("GOOD " + string_good_data)
  print("BAD" + string_failed_results)

  if(len(string_failed_results) > 10):
    full_prompt = clean_prompt + string_good_data + "\n\nHere is the JSON structure for expected formatting:\n" + json_string + "\n\nList to correct:\n" + string_failed_results


    response = client.chat.completions.create(
                model="Llama-3.3-Swallow-70B-Instruct-v0.4",
                messages=[
                    {
                        "role": "user",
                        "content": full_prompt
                    }
                ]
            )
    
    response_content = response.choices[0].message.content.strip()

    print("\n"*10)
    print(response_content)
    return response_content
