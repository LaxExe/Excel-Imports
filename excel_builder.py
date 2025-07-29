from openpyxl import Workbook
import json
from row_parsing import gather_row_data

def export_to_excel(data, output_file):
    """
    Export parsed data to Excel.
    Ensures:
    - Phone numbers are saved as strings to avoid Excel auto-formatting
    - Long numbers (e.g. credit card-like) are not turned into scientific notation
    - Headers match dictionary keys from the first row
    """
    # Check if we have data to export
    if not data:
        print("No data to write.")
        return

    # Create new Excel workbook and get the active worksheet
    wb = Workbook()
    ws = wb.active

    # Get headers from the first row of data
    # Headers will be the keys from the dictionary (email, phone_number, full_name, etc.)
    headers = list(data[0].keys())
    ws.append(headers)

    # Process each row of data
    for row_dict in data:
        row = []
        for header in headers:
            # Get the value for this column, default to empty string if missing
            val = row_dict.get(header, "")

            # Always convert to string to preserve exact formatting
            # This prevents Excel from auto-formatting our data
            val = str(val)

            # Special handling for phone numbers
            # Phone numbers need special care to avoid Excel's auto-formatting
            if header.lower() == "phone_number":
                # Don't add apostrophe for phone numbers as they're already properly formatted
                # Just ensure they're treated as text
                if val and val != "INVALID":
                    # For phone numbers with extensions, ensure proper formatting
                    if " x" in val:
                        # Already properly formatted with extension (e.g., +16135551212 x123)
                        # No need to modify, just keep as is
                        pass
                    elif val.startswith("+"):
                        # International format (e.g., +16135551212), keep as is
                        # The + prefix ensures Excel treats it as text
                        pass
                    else:
                        # Regular phone number (e.g., 16135551212), ensure it's treated as text
                        # Excel might try to format this as a number, but we want to preserve it
                        pass
            elif val.replace(" ", "").replace("-", "").replace(".", "").replace("+", "").isdigit() and len(val) > 10:
                # Only add apostrophe for very long numeric values that might be scientific notation
                # This prevents Excel from converting large numbers to scientific notation
                # Examples: credit card numbers, long IDs, etc.
                val = f"'{val}"

            # Add the processed value to the row
            row.append(val)
        
        # Add the complete row to the worksheet
        ws.append(row)

    # Save the workbook to the specified file
    wb.save(output_file)
    print(f"Excel file saved as: {output_file}")



