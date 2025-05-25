import pandas as pd
from faker import Faker
import random
from datetime import datetime, timedelta

# Initialize Faker with a seed for reproducibility
fake = Faker('en_US')
Faker.seed(42)

# Define constants
NUM_ROWS = 20
FLORIDA_CITIES = [
    ('Miami', '33186'), ('Key Biscayne', '33149'), ('Coral Gables', '33143'), 
    ('Miami Lakes', '33014'), ('Miami Gardens', '33055')
]
SUBMISSION_START_DATE = datetime(2025, 4, 1)
SUBMISSION_END_DATE = datetime(2025, 5, 16)

# Function to generate a random date in MMM DD, YYYY format
def random_date(start_date, end_date):
    delta = end_date - start_date
    random_days = random.randint(0, delta.days)
    return (start_date + timedelta(days=random_days)).strftime('%b %d, %Y')

# Function to generate a random DOB between 1950 and 2020
def random_dob():
    start_date = datetime(1950, 1, 1)
    end_date = datetime(2020, 12, 31)
    return random_date(start_date, end_date)

# Generate fake data
data = []
for _ in range(NUM_ROWS):
    # Common fields
    city, zip_code = random.choice(FLORIDA_CITIES)
    row = {
        'Submission Date': random_date(SUBMISSION_START_DATE, SUBMISSION_END_DATE),
        'Street Address': fake.street_address(),
        'Apartment Number': fake.secondary_address() if random.random() > 0.5 else '',
        'City': city,
        'State': 'Florida',
        'Zip Code': zip_code,
        'Primary Telephone Number': fake.phone_number(),
        'Primary Student E-mail': fake.email(),
        'First Name': fake.first_name(),
        'Middle Name': fake.first_name() if random.random() > 0.3 else '',
        'Last Name': fake.last_name(),
        'First Student Birth Date': random_dob(),
        'First Student Gender': random.choice(['Male', 'Female']),
        'First Name_2': '',  # Second student First Name
        'Middle Name_2': '',
        'Last Name_2': '',
        'Second Student Birth Date': '',
        'Second Student Gender': '',
        'First Name_3': '',  # Third student First Name
        'Middle Name_3': '',
        'Last Name_3': '',
        'Third Student Birth Date': '',
        'Third Student Gender': '',
        'First Name_4': '',  # Fourth student First Name
        'Middle Name_4': '',
        'Last Name_4': '',
        'Fourth Student Birth Date': '',
        'Fourth Student Gender': ''
    }
    
    # Second student (~50% chance)
    if random.random() > 0.5:
        row['First Name_2'] = fake.first_name()
        row['Middle Name_2'] = fake.first_name() if random.random() > 0.3 else ''
        row['Last Name_2'] = fake.last_name()
        row['Second Student Birth Date'] = random_dob()
        row['Second Student Gender'] = random.choice(['Male', 'Female'])
    
    # Third student (~30% chance)
    if random.random() > 0.7:
        row['First Name_3'] = fake.first_name()
        row['Middle Name_3'] = fake.first_name() if random.random() > 0.3 else ''
        row['Last Name_3'] = fake.last_name()
        row['Third Student Birth Date'] = random_dob()
        row['Third Student Gender'] = random.choice(['Male', 'Female'])
    
    # Fourth student (~20% chance)
    if random.random() > 0.8:
        row['First Name_4'] = fake.first_name()
        row['Middle Name_4'] = fake.first_name() if random.random() > 0.3 else ''
        row['Last Name_4'] = fake.last_name()
        row['Fourth Student Birth Date'] = random_dob()
        row['Fourth Student Gender'] = random.choice(['Male', 'Female'])
    
    data.append(row)

# Create DataFrame with duplicate column names
columns = [
    'Submission Date', 'Street Address', 'Apartment Number', 'City', 'State', 'Zip Code',
    'Primary Telephone Number', 'Primary Student E-mail',
    'First Name', 'Middle Name', 'Last Name', 'First Student Birth Date', 'First Student Gender',
    'First Name', 'Middle Name', 'Last Name', 'Second Student Birth Date', 'Second Student Gender',
    'First Name', 'Middle Name', 'Last Name', 'Third Student Birth Date', 'Third Student Gender',
    'First Name', 'Middle Name', 'Last Name', 'Fourth Student Birth Date', 'Fourth Student Gender'
]
# Use internal keys to avoid duplicate column issues in DataFrame
internal_columns = [
    'Submission Date', 'Street Address', 'Apartment Number', 'City', 'State', 'Zip Code',
    'Primary Telephone Number', 'Primary Student E-mail',
    'First Name', 'Middle Name', 'Last Name', 'First Student Birth Date', 'First Student Gender',
    'First Name_2', 'Middle Name_2', 'Last Name_2', 'Second Student Birth Date', 'Second Student Gender',
    'First Name_3', 'Middle Name_3', 'Last Name_3', 'Third Student Birth Date', 'Third Student Gender',
    'First Name_4', 'Middle Name_4', 'Last Name_4', 'Fourth Student Birth Date', 'Fourth Student Gender'
]
df = pd.DataFrame(data, columns=internal_columns)

# Write to Excel with desired column names (including duplicates)
with pd.ExcelWriter('Fake_KBYC_Source_2025.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Sheet1', columns=internal_columns)
    # Rename columns in the Excel file to match the desired duplicates
    worksheet = writer.sheets['Sheet1']
    for idx, col_name in enumerate(columns, start=1):
        worksheet.cell(row=1, column=idx).value = col_name

print(f"Fake source file generated: Fake_KBYC_Source_2025.xlsx")