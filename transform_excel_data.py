import pandas as pd
import warnings

# Suppress warnings for cleaner output
warnings.filterwarnings("ignore")

# Read the source Excel file
source_file = "May_1_2025_-_KBYC_-_Boat_Americ2025-05-17_07_57_29.xlsx"
source_df = pd.read_excel(source_file, sheet_name="Sheet1")

# Normalize column names: strip whitespace
source_df.columns = [col.strip() for col in source_df.columns]

# Log all column names for debugging
print("Source DataFrame Columns:", list(source_df.columns))

# Define the target columns based on FWC Results structure
target_columns = [
    "Provider ID", "Test Date", "First Name", "MI", "Last Name",
    "Street Address", "Apartment # (address continued)", "City", "State",
    "County", "Country Code", "Zip", "Zip Extn", "DOB", "Test Version",
    "Violation Req", "Proctored", "Pass", "Gender", "E-card Y/N",
    "E-mail Address", "Re-try E-mail Address"
]

# Initialize a list to store transformed rows
transformed_rows = []

# Process each row in the source DataFrame
for _, row in source_df.iterrows():
    # Define student sets (First, Second, Third, Fourth)
    student_sets = [
        ("First", "First ", ""),  # First student uses base column names
        ("Second", "Second ", ".1"),
        ("Third", "Third ", ".2"),
        ("Fourth", "Fourth ", ".3")
    ]
    
    # Extract common address fields
    address = str(row["Street Address"]) if pd.notna(row["Street Address"]) else ""
    apt = str(row["Apartment Number"]) if pd.notna(row["Apartment Number"]) else ""
    city = str(row["City"]) if pd.notna(row["City"]) else ""
    state = str(row["State"]) if pd.notna(row["State"]) else ""
    zip_code = str(row["Zip Code"]) if pd.notna(row["Zip Code"]) else ""
    
    # Use Primary Student E-mail as default
    email = str(row["Primary Student E-mail"]) if pd.notna(row["Primary Student E-mail"]) else ""
    
    # Process each student in the row
    for student_label, dob_gender_prefix, name_suffix in student_sets:
        # Access student fields with correct column names
        first_name_key = f"First Name{name_suffix}" if name_suffix else "First Name"
        middle_name_key = f"Middle Name{name_suffix}" if name_suffix else "Middle Name"
        last_name_key = f"Last Name{name_suffix}" if name_suffix else "Last Name"
        dob_key = f"{dob_gender_prefix}Student Birth Date"
        gender_key = f"{dob_gender_prefix}Student Gender"
        
        first_name = row[first_name_key] if pd.notna(row.get(first_name_key)) else None
        middle_name = row[middle_name_key] if pd.notna(row.get(middle_name_key)) else ""
        last_name = row[last_name_key] if pd.notna(row.get(last_name_key)) else None
        dob = row[dob_key] if pd.notna(row.get(dob_key)) else None
        gender = row[gender_key] if pd.notna(row.get(gender_key)) else None
        
        # Debug: Log student data extraction
        print(f"Processing {student_label} Student: First Name={first_name}, Last Name={last_name}, DOB={dob}, Gender={gender}")
        
        # Only create a row if first_name and last_name are present
        if first_name and last_name and pd.notna(first_name) and pd.notna(last_name):
            # Format DOB to match target (MM/DD/YYYY)
            dob_str = ""
            if pd.notna(dob):
                try:
                    # Convert DOB to string to handle various input types
                    dob_input = str(dob)
                    # Parse DOB, allowing pandas to infer format
                    dob_parsed = pd.to_datetime(dob_input, errors="coerce")
                    if pd.notna(dob_parsed):
                        dob_str = dob_parsed.strftime("%m/%d/%Y")
                    else:
                        print(f"Warning: Failed to parse DOB '{dob_input}' (raw: {dob}, type: {type(dob)}) for {first_name} {last_name}")
                except Exception as e:
                    print(f"Error parsing DOB '{dob_input}' (raw: {dob}, type: {type(dob)}) for {first_name} {last_name}: {e}")
            
            # Create the transformed row
            transformed_row = {
                "Provider ID": "2",  # From sample
                "Test Date": "5/17/25",  # From sample
                "First Name": first_name,
                "MI": middle_name[:1] if middle_name and pd.notna(middle_name) else "",  # Take first letter of middle name
                "Last Name": last_name,
                "Street Address": address,
                "Apartment # (address continued)": apt,
                "City": city,
                "State": state,
                "County": "Miami-Dade",  # Assumed based on locations
                "Country Code": "US",  # Default
                "Zip": zip_code,
                "Zip Extn": "",  # Not provided
                "DOB": dob_str,
                "Test Version": "G",  # From sample
                "Violation Req": "N",  # From sample
                "Proctored": "Y",  # From sample
                "Pass": "Y",  # From sample
                "Gender": gender if gender and pd.notna(gender) else "",
                "E-card Y/N": "Y",  # From sample
                "E-mail Address": email,  # Use primary email
                "Re-try E-mail Address": email,  # Use primary email
            }
            transformed_rows.append(transformed_row)
            print(f"Added row for {first_name} {last_name}")

# Create the target DataFrame
target_df = pd.DataFrame(transformed_rows, columns=target_columns)

# Write to a new Excel file
output_file = "FWC_Results_Transformed_05172025.xlsx"
target_df.to_excel(output_file, index=False, sheet_name="Sheet1")

print(f"Transformation complete. Output saved to {output_file}")