# Identify type of address using the json
import json


def identify_type(filepath):

    with open(filepath, "r") as f:
        info_json = json.load(f)

    is_1_column = info_json["required_fields"]["address_takes_up_1_column"]

    return is_1_column




def column_1_address(address, format_str, separator):
    format_parts = [f.strip() for f in format_str.split(separator)]
    address_parts = [a.strip() for a in address.split(separator)]

    result = {}

    extra_count = len(address_parts) - len(format_parts)


    if extra_count > 0:
            leftover = address_parts[:extra_count]
            street_value = separator.join(leftover).strip()
    else:
        #Call on AI
        print("not enough parts")
        street_value = "missing"

    format_without_street = [part for part in format_parts if part != 'street_address']

    addr_index = extra_count
    for part in format_without_street:
        if addr_index < 0 or addr_index >= len(address_parts):
            result[part] = "missing"
            #call on ai
        else:
            result[part] = address_parts[addr_index]
        addr_index += 1
        
    result['street_address'] = street_value

    return result

print(column_1_address("123 Main St, Toronto, ON, M4C 1A1, Canada","street_address, city, province, postal_code, country", "," ))





# Case1: Address takes up 1 column
    # Using seporators arrange address with address info form json into Unit, Street, City, Province, Postal, Country. make use of seperator and map in reverse
    # use methods for missing info
    # case 1, one of the feilds are missing pass to ai
        # return pass
    #return Unit, Street, City, Province, Postal, Country if anything is missing fill out address with missing info





