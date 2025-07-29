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
    Extracts a region code (e.g., 'US', 'CA') from an address string.
    First checks common aliases like "USA", "UK", "UAE", then tries pycountry.
    """
    if not address:
        return None

    address = address.lower()

    # Manual alias map to catch common variations
    alias_map = {
        "usa": "US",
        "united states": "US",
        "us": "US",
        "uk": "GB",
        "uae": "AE",
        "canada": "CA",
        "india": "IN"
    }

    for alias, code in alias_map.items():
        if alias in address:
            return code

    # Try matching full country names using pycountry
    for country in pycountry.countries:
        if country.name.lower() in address:
            return country.alpha_2

    return None


def clean_phone_number(raw_phone, full_name, email, address):
    """
    Cleans and standardizes a phone number:
    - Preserves extensions like 'x123'
    - Converts numbers like 1.234E+12 from Excel to valid format
    - Uses address to infer region if country code is missing
    - If valid, returns number in +E.164 format (e.g., +16135551212)
    - If no region can be inferred, keeps raw cleaned number
    - If completely invalid, returns 'INVALID'
    """

    # If phone is empty but we have name and email, return empty string
    # This handles cases where phone is optional but other fields are present
    if (not raw_phone or str(raw_phone).strip() == "") and full_name.strip() and email.strip():
        return ""

    # Convert to string and strip whitespace for consistent processing
    raw_str = str(raw_phone).strip()

    # Handle Excel scientific notation (e.g., 1.234E+12)
    # Excel sometimes converts large numbers to scientific notation, we need to convert back
    if 'E' in raw_str.upper() or 'e' in raw_str:
        try:
            # Convert scientific notation to regular number
            # This handles cases like 1.234E+12 becoming 1234000000000
            raw_str = str(int(float(raw_str)))
        except (ValueError, OverflowError):
            # If conversion fails, the number is invalid
            return "INVALID"

    # Extract extension like x123 or ext. 456
    # Extensions are important and should be preserved
    # We look for 'x' or 'ext' followed by numbers
    ext_match = re.search(r"(x|ext\.?)\s*(\d+)", raw_str, re.IGNORECASE)
    extension = f" x{ext_match.group(2)}" if ext_match else ""

    # Remove extension from main phone string for cleaning
    # We clean the main number separately, then add extension back
    if ext_match:
        raw_str = raw_str[:ext_match.start()].strip()

    # Clean main part (preserve + if present)
    # Remove all non-digit characters except + (which indicates country code)
    phone_str = re.sub(r"[^\d+]", "", raw_str)

    # Handle US country code formatting
    # This is the key logic for standardizing US phone numbers
    if phone_str.startswith('+'):
        # Already has +, keep as is (e.g., +17837799889)
        pass
    elif phone_str.startswith('001'):
        # Convert 001 to +1 (international dialing code to E.164 format)
        # 001 is the international dialing code for US/Canada, equivalent to +1
        phone_str = '+1' + phone_str[3:]
    elif phone_str.startswith('1') and len(phone_str) >= 11:
        # If it starts with 1 and is long enough to be a US number, add +
        # This handles cases like 19668270528 becoming +19668270528
        phone_str = '+' + phone_str
    else:
        # Remove leading zeros for other numbers (non-US numbers)
        # Leading zeros are not valid in international format
        phone_str = phone_str.lstrip('0')

    # If phone is too short after cleaning, it's invalid
    # Minimum length for a valid phone number is 7 digits
    if len(phone_str) < 7:
        return "INVALID"

    # Extract country code from address to help with validation
    # This helps determine if it's a US number for better validation
    region = extract_country_code(address)

    try:
        # Try parsing with region or default to None
        # Use the phonenumbers library for proper validation
        parsed = phonenumbers.parse(phone_str, region or None)
        if phonenumbers.is_valid_number(parsed):
            # If valid, format in E.164 international format
            # E.164 format: +[country code][number] (e.g., +16135551212)
            formatted = phonenumbers.format_number(parsed, phonenumbers.PhoneNumberFormat.E164)
            return formatted + extension
        else:
            # If not valid but has reasonable length, return cleaned version
            # Sometimes the library is too strict, so we check length manually
            digits_only = phone_str.replace('+', '')
            if 10 <= len(digits_only) <= 15:
                # Standard phone number length range
                return phone_str + extension
            # For US numbers, if it's 9 digits, it might be missing area code
            # Some valid US numbers might be missing the area code
            elif len(digits_only) == 9 and region == "US":
                # Could be a valid number missing area code, return as is
                return phone_str + extension
            return "INVALID"
    except phonenumbers.NumberParseException:
        # Last resort: if it's all digits and 10–15 digits, treat as raw
        # If the library can't parse it, we do basic validation ourselves
        digits_only = phone_str.replace('+', '')
        if re.fullmatch(r"\d{10,15}", digits_only):
            # Standard phone number length, accept it
            return phone_str + extension
        # For US numbers, also accept 9 digits as they might be valid
        # Some US numbers might be missing area code but still valid
        elif re.fullmatch(r"\d{9}", digits_only) and region == "US":
            return phone_str + extension
        return "INVALID"


def clean_ai_response(response_text: str) -> str:
    text = response_text.strip()
    if text.startswith("```"):
        first_newline = text.find("\n")
        if first_newline != -1:
            text = text[first_newline + 1:]
        if text.endswith("```"):
            text = text[:-3]
    return text.strip()


clean_prompt = """you will be given a list of information that did not meet the cirteria, it may be because there are additional spaces, commas or some incorect charecters,
                 you will be given a json of how the infomration was supposed to be arranged and a set of data with proper information, your job will be to fix up the list 
                 and return only the list with no formating or any other informatin formated or explanation provided, just return a list with the corrected information,
                 the list without mistakes is below, then a list with mistakes is after that"""

def AI_check(results, failed_results):
    
  string_good_data = " ".join(results)
  string_failed_results = " ".join(failed_results)

  print(string_good_data)
  print(string_failed_results)

  response = client.chat.completions.create(
              model="Llama-3.3-Swallow-70B-Instruct-v0.4",
              messages=[
                  {
                      "role": "user",
                      "content": clean_prompt + string_good_data + json_string +" list that you need to fix : " + string_failed_results
                  }
              ]
          )
  
  response_content = response.choices[0].message.content.strip()

  print("\n"*10)
  print(response_content)
  return response_content
