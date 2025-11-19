import os
import re
from datetime import datetime

# Lookup table for phone numbers and their corresponding file types
phone_number_lookup = {
    "+18666223696": "Re Auth",
    # Add more phone numbers and types as needed
}

# Directory containing the files (change this if needed)
directory = "."

# Pattern to match filenames
pattern = re.compile(r"(\+\d+)-(\d{4})-(\d{6})-(\d+)\.pdf")

# Track counters for each base filename to avoid overwriting
counters = {}

# Iterate through files in the directory
for filename in os.listdir(directory):
    match = pattern.match(filename)
    if match:
        phone_number = match.group(1)
        mmdd = match.group(2)

        if phone_number in phone_number_lookup:
            file_type = phone_number_lookup[phone_number]
            base_name = f"2025{mmdd} {file_type}"

            # Initialize or increment the counter
            counters.setdefault(base_name, 0)
            counters[base_name] += 1

            new_filename = f"{base_name} {counters[base_name]}.pdf"
            old_path = os.path.join(directory, filename)
            new_path = os.path.join(directory, new_filename)

            os.rename(old_path, new_path)
            print(f"Renamed '{filename}' to '{new_filename}'")
        else:
            print(f"No match for phone number in '{filename}', skipping.")
    else:
        print(f"Filename '{filename}' does not match expected pattern, skipping.")
