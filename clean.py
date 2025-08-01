import re
import phonenumbers
import pycountry
from dotenv import load_dotenv
import os
from openai import OpenAI
import json

load_dotenv()
api_key = os.getenv("OPENAI_API_KEY")

client = OpenAI(
    api_key=api_key,
    base_url="https://api.sambanova.ai/v1"
)




# Validation Function
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
  
  return name


def validate_email(email):
  if not email or "@" not in email:
    return "MISSING"
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
    elif phone_str.startswith('1') and len(phone_str) >= 11:        # If it starts with 1 and is long enough to be a US number, add +

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

    return "MISSING"

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

Be precise and conservative. Only clean and rearrange as needed to make them structurally identical to the examples in the good list, if there are elements or portions of elements just add a null value for them.

CRITICAL: Return ONLY valid JSON with no extra text, explanations, or formatting. you have flexibility over the address components but no other feilds, maining if there is a country within in the good data you can add a country feild, if there is no provice or state name within the good data you can remove it.

{
  "clean_items": [
    {
      "missing_parts_of_address": true,
      "email": "valid_email",
      "phone_number": "valid_phone_number",
      "full_name": "validate_name",
      "address": [
        {
          "street_address": "string",
          "postal_code": "string or null",
          "city": "string or null",
          "province_or_state_name":  Do not include this feild in the data if it is not present in the good data examples,
        }
      ],
      "additional_fields": "valid_items"
    }
  ]
}

The list of correct examples is:
"""

def AI_check(results, failed_results, page):
    
  string_good_data = str(results)
  string_failed_results = str(failed_results)

  print("GOOD " + string_good_data)
  print("BAD" + string_failed_results)

  if(len(string_failed_results) > 10):
    full_prompt = clean_prompt + string_good_data + "\n\nList to correct:\n" + string_failed_results

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

    try:
        cleaned_response = clean_ai_response(response_content)
        
        if not cleaned_response.startswith('{'):
            start = cleaned_response.find('{')
            end = cleaned_response.rfind('}') + 1
            if start != -1 and end != 0:
                cleaned_response = cleaned_response[start:end]
        
        remake_json = json.loads(cleaned_response)
        required_fields = ["clean_items"]
        all_required_present = True

        for field in required_fields:
            value = remake_json.get(field)
            if value is None or value == "" or value == "null":
                print(f"Missing required field '{field}' in the response.")
                all_required_present = False
                break
            
            if field == "clean_items":
                if not isinstance(value, list) or len(value) == 0:
                    print(f"Field '{field}' must be a non-empty list.")
                    all_required_present = False
                    break

        if all_required_present:
            with open(f"remake{page}.json", "w") as f:
                json.dump(remake_json, f, indent=4) 

    except json.JSONDecodeError:
        print("Error decoding JSON response from OpenAI API.")
        pass

    print(cleaned_response)


    

    return "remake.json"

