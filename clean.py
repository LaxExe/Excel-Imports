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
