# Excel Data Transformation Script

## Overview
This Python script transforms student data from an Excel file (e.g., `May_1_2025_-_KBYC_-_Boat_Americ2025-05-17_07_57_29.xlsx` or `Fake_KBYC_Source_2025.xlsx`) to match the structure of `FWC Results 05172025.xls`. It processes multiple students per row, mapping fields to the target format and filling in default values where necessary. The project is designed to help manage the Public Education registration of the Boating Class for the USCG Auxiliary of Flotilla 65, facilitating efficient processing of student registrations for boating safety courses (see [USCG Auxiliary Class Finder](https://cgaux.org/boatinged/class_finder/)). A utility script (`generate_fake_source.py`) creates a fake source Excel file for testing, using duplicate column names (`First Name`, `Middle Name`, `Last Name`) for each student, which pandas reads with suffixes (e.g., `First Name.1`).

## Prerequisites
- **Python 3.x** installed.
- **UV** (recommended package manager):
  - macOS and Linux: `curl -LsSf https://astral.sh/uv/install.sh | sh`
  - Windows (PowerShell, admin): `powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"`
  - Homebrew: `brew install uv`
- Required Python libraries:
  - `pandas`
  - `openpyxl`
  - `faker` (for generating fake data)
- Install dependencies using:
  ```bash
  uv add pandas openpyxl faker
  ```

## Usage
1. **Place Files**:
   - Ensure the source Excel file (e.g., `May_1_2025_-_KBYC_-_Boat_Americ2025-05-17_07_57_29.xlsx` or `Fake_KBYC_Source_2025.xlsx`) is in the same directory as the transformation script (`transform_excel_data.py`).
2. **Generate Fake Source File (Optional)**:
   - To create a fake source file for testing:
     ```bash
     uv run generate_fake_source.py
     ```
     This generates `Fake_KBYC_Source_2025.xlsx` with ~20 rows of fake student data.
   - Update the `source_file` variable in `transform_excel_data.py` to use the fake file:
     ```python
     source_file = "Fake_KBYC_Source_2025.xlsx"
     ```
3. **Run the Transformation Script**:
   - Execute the transformation script:
     ```bash
     uv run transform_excel_data.py
     ```
4. **Output**:
   - The script generates `FWC_Results_Transformed_05172025.xlsx` containing the transformed data.

## Input and Output
- **Input File**:
  - Original: `May_1_2025_-_KBYC_-_Boat_Americ2025-05-17_07_57_29.xlsx`
  - Fake (for testing): `Fake_KBYC_Source_2025.xlsx`
  - Structure:
    - Columns: `Submission Date`, `Street Address`, `Apartment Number`, `City`, `State`, `Zip Code`, `Primary Telephone Number`, `Primary Student E-mail`
    - First student: `First Name`, `Middle Name`, `Last Name`, `First Student Birth Date`, `First Student Gender`
    - Second student: `First Name`, `Middle Name`, `Last Name`, `Second Student Birth Date`, `Second Student Gender`
    - Third student: `First Name`, `Middle Name`, `Last Name`, `Third Student Birth Date`, `Third Student Gender`
    - Fourth student: `First Name`, `Middle Name`, `Last Name`, `Fourth Student Birth Date`, `Fourth Student Gender`
    - Note: Duplicate column names (`First Name`, etc.) are read by pandas as `First Name`, `First Name.1`, `First Name.2`, `First Name.3`, etc.
- **Output File**: `FWC_Results_Transformed_05172025.xlsx`
  - Matches `FWC Results 05172025.xls` structure with columns: `Provider ID`, `Test Date`, `First Name`, `MI`, `Last Name`, `Street Address`, `Apartment # (address continued)`, `City`, `State`, `County`, `Country Code`, `Zip`, `Zip Extn`, `DOB`, `Test Version`, `Violation Req`, `Proctored`, `Pass`, `Gender`, `E-card Y/N`, `E-mail Address`, `Re-try E-mail Address`.
  - Each student is represented as a separate row, resulting in ~30–40 rows for typical input files.

## Script Details
- **Transformation Logic**:
  - Maps source fields to target fields (e.g., `First Name` to `First Name`, `First Student Birth Date` to `DOB` formatted as `MM/DD/YYYY`).
  - Handles multiple students per row (First, Second, Third, Fourth) by creating separate rows for each valid student (with non-empty `First Name` and `Last Name`).
  - Sets default values for fields not in the source, based on `FWC Results 05172025.xls`:
    - `Provider ID`: `2`
    - `Test Date`: `5/17/25`
    - `Test Version`: `G`
    - `Violation Req`: `N`
    - `Proctored`: `Y`
    - `Pass`: `Y`
    - `E-card Y/N`: `Y`
    - `County`: `Miami-Dade` (inferred from locations)
    - `Country Code`: `US`
  - Uses `Primary Student E-mail` for all students in a row, as individual emails are not provided.
  - Leaves optional fields like `Zip Extn` and `Re-try E-mail Address` to empty.
- **Dependencies**:
  - `pandas`: For Excel file reading and data manipulation.
  - `openpyxl`: For reading and writing Excel files.
  - `faker`: For generating fake test data (used in `generate_fake_source.py`).
- **Column Sensitivity**:
  - The transformation script expects `First Name`, `First Name.1`, `First Name.2`, `First Name.3` due to pandas’ handling of duplicate column names in the source file.

## Testing
- **Using Fake Data**:
  - Generate a fake source file with:
    ```bash
    uv run generate_fake_source.py
    ```
  - Ensure `transform_excel_data.py` uses `source_file = "Fake_KBYC_Source_2025.xlsx"`.
  - Run the transformation script:
    ```bash
    uv run transform_excel_data.py
    ```
  - Verify the output (`FWC_Results_Transformed_05172025.xlsx`):
    - ~30–40 rows (depending on Second, Third, Fourth students).
    - DOBs in `MM/DD/YYYY` format (