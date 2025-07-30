import openpyxl
import json
import re
import pycountry
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderUnavailable, GeocoderServiceError


geolocator = Nominatim(user_agent="excel_imports_app")

def street_and_postal_code(street, postal_code):
 
  if not street and not postal_code:
    return "Street and postal code are missing"
 
  try:
    if street and postal_code:
      location = geolocator.geocode(f"{street} {postal_code}", exactly_one=True, timeout=10)
      return location if location else "Unable to extract address"
    else:
      return "No Results Found"
  except GeocoderTimedOut:
    return "No Results Found"
  except GeocoderServiceError as e:
    return "No Results Found"
  except Exception as e:
    return "No Results Found"
 

streets = "585 Twain Avenue"
postal_codes = "L5W1M1"

address = street_and_postal_code(streets, postal_codes)
raw_address = address.raw

import json
print(json.dumps(raw_address, indent=2))
