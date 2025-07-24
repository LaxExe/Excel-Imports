import re
import phonenumbers
import pycountry


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



