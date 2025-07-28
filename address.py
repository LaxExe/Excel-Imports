import json


def is_1_column_tag(filepath):

    with open(filepath, "r") as f:
        info_json = json.load(f)

    is_1_column = info_json["required_fields"]["address_takes_up_1_column"]

    return is_1_column

def result_format(result, required_format, skip):
    if skip:
        return True
    
    formatted_parts = []

    for key in required_format:
        if key in result:
            formatted_parts.append(result[key])
        else:
            return True

    formatted_string = ", ".join(formatted_parts)
    return formatted_string


def column_1_address_skip(address, format_str, separator):
    format_parts = [f.strip() for f in format_str.split(separator)]
    address_parts = [a.strip() for a in address.split(separator)]

    required_format = [ "Postal Code", "Street", "City", "Province"]

    result = {}
    skip = False

    extra_count = len(address_parts) - len(format_parts)

    if extra_count >= 0:
        for i, key in enumerate(format_parts):
            if i == 0:
                value = separator.join(address_parts[:extra_count + 1]).strip()
            else:
                value = address_parts[i + extra_count]

            if len(value) < 2 or len(value) > 40:
                skip = True  
                break

            result[key] = value
    else:
        skip = True

    if skip:
        return skip

    result_string = result_format(result, required_format, skip)

    return result_string







   


