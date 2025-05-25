# Flotilla65FWCBoatingEdTransformer

## Overview
The **Flotilla65FWCBoatingEdTransformer** is a Python-based tool designed to streamline the Public Education registration process for the Boating Class offered by the USCG Auxiliary Flotilla 65. It transforms student registration data from an Excel file (e.g., `May_1_2025_-_KBYC_-_Boat_Americ2025-05-17_07_57_29.xlsx`) into a standardized format matching `FWC Results 05172025.xls`, facilitating efficient management of boating safety course registrations. The project includes a transformation script (`transform_excel_data.py`) and a utility script (`generate_fake_source.py`) for creating fake test data. For more information on Flotilla 65’s boating classes, visit the [USCG Auxiliary Class Finder](https://cgaux.org/boatinged/class_finder/).

The source Excel files use duplicate column names (`First Name`, `Middle Name`, `Last Name`) for each student (First, Second, Third, Fourth), which pandas reads with suffixes (e.g., `First Name.1`). This tool handles such structures to produce a clean, per-student output.

## Project Structure
- `transform_excel_data.py`: Main script to transform student registration data into the target FWC format.
- `generate_fake_source.py`: Utility script to generate a fake source Excel file (`Fake_KBYC_Source_2025.xlsx`) for testing.
- `README.md`: Project documentation (this file).
- `LICENSE`: MIT License governing the project.
- `pyproject.toml`: Project metadata and dependency configuration.
- `uv.lock`: Lock file for reproducible dependency installation using `uv`.

## Prerequisites
- **Python 3.12 or higher** (as specified in `pyproject.toml`).
- **UV** (recommended package manager for dependency management):
  - macOS/Linux: `curl -LsSf https://astral.sh/uv/install.sh | sh`
  - Windows (PowerShell, admin): `powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"`
  - Homebrew: `brew install uv`
- Required Python libraries (specified in `pyproject.toml`):
  - `pandas>=2.2.3`
  - `openpyxl>=3.1.5`
  - `faker>=37.3.0`
- Install dependencies using:
  ```bash
  uv sync
  ```

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/jjborie/Flotilla65FWCBoatingEdTransformer.git
   cd Flotilla65FWCBoatingEdTransformer
   ```
2. Set up the project with `uv`:
   ```bash
   uv sync
   ```
   This installs `pandas`, `openpyxl`, `faker`, and their dependencies as specified in `pyproject.toml` and `uv.lock`.

## Usage
1. **Prepare the Source File**:
   - Place the source Excel file (e.g., `May_1_2025_-_KBYC_-_Boat_Americ2025-05-17_07_57_29.xlsx`) in the project directory.
   - Alternatively, generate a fake source file for testing (see below).

2. **Generate Fake Source File (Optional)**:
   - Run the fake data generator:
     ```bash
     uv run generate_fake_source.py
     ```
     This creates `Fake_KBYC_Source_2025.xlsx` with ~20 rows of fake student data.
   - Update `transform_excel_data.py` to use the fake file by setting:
     ```python
     source_file = "Fake_KBYC_Source_2025.xlsx"
     ```

3. **Run the Transformation Script**:
   - Execute the transformation:
     ```bash
     uv run transform_excel_data.py
     ```
   - Output: `FWC_Results_Transformed_05172025.xlsx` containing the transformed data.

## Input and Output
- **Input File**:
  - Original: `May_1_2025_-_KBYC_-_Boat_Americ2025-05-17_07_57_29.xlsx`
  - Fake (for testing): `Fake_KBYC_Source_2025.xlsx`
  - Structure:
    - Columns:
      ```
      Submission Date, Street Address, Apartment Number, City, State, Zip Code,
      Primary Telephone Number, Primary Student E-mail,
      First Name, Middle Name, Last Name, First Student Birth Date, First Student Gender,
      First Name, Middle Name, Last Name, Second Student Birth Date, Second Student Gender,
      First Name, Middle Name, Last Name, Third Student Birth Date, Third Student Gender,
      First Name, Middle Name, Last Name, Fourth Student Birth Date, Fourth Student Gender
      ```
    - Note: Duplicate `First Name`, `Middle Name`, `Last Name` columns are read by pandas as `First Name`, `First Name.1`, `First Name.2`, `First Name.3`, etc.
    - Data: Each row contains up to four students, with ~50% chance for Second, ~30% for Third, and ~20% for Fourth students in fake data.

- **Output File**: `FWC_Results_Transformed_05172025.xlsx`
  - Structure: Matches `FWC Results 05172025.xls` with columns:
    ```
    Provider ID, Test Date, First Name, MI, Last Name, Street Address,
    Apartment # (address continued), City, State, County, Country Code,
    Zip, Zip Extn, DOB, Test Version, Violation Req, Proctored, Pass,
    Gender, E-card Y/N, E-mail Address, Re-try E-mail Address
    ```
  - Each student is a separate row, yielding ~30–40 rows for typical input files.
  - DOBs are formatted as `MM/DD/YYYY` (e.g., `12/23/1953`).

## Script Details
- **Transformation Logic** (`transform_excel_data.py`):
  - Maps source fields to target fields (e.g., `First Name.1` to `First Name` for Second student).
  - Processes multiple students per row, creating separate rows for each valid student (non-empty `First Name` and `Last Name`).
  - Applies default values:
    - `Provider ID`: `2`
    - `Test Date`: `5/17/25`
    - `Test Version`: `G`
    - `Violation Req`: `N`
    - `Proctored`: `Y`
    - `Pass`: `Y`
    - `E-card Y/N`: `Y`
    - `County`: `Miami-Dade`
    - `Country Code`: `US`
  - Uses `Primary Student E-mail` for both `E-mail Address` and `Re-try E-mail Address`.
  - Leaves `Zip Extn` empty.
- **Fake Data Generation** (`generate_fake_source.py`):
  - Creates ~20 rows with realistic Florida addresses, emails, and student data.
  - Uses duplicate column names (`First Name`, etc.), which pandas reads with suffixes.
- **Dependencies** (from `pyproject.toml`):
  - `pandas>=2.2.3`: Excel file reading and data manipulation.
  - `openpyxl>=3.1.5`: Excel file reading and writing.
  - `faker>=37.3.0`: Fake data generation.
- **Column Sensitivity**:
  - The transformation script handles `First Name`, `First Name.1`, `First Name.2`, `First Name.3` due to pandas’ duplicate column handling.

## Testing
- **Using Fake Data**:
  - Generate fake data:
    ```bash
    uv run generate_fake_source.py
    ```
  - Set `source_file = "Fake_KBYC_Source_2025.xlsx"` in `transform_excel_data.py`.
  - Run transformation:
    ```bash
    uv run transform_excel_data.py
    ```
  - Verify `FWC_Results_Transformed_05172025.xlsx`:
    - ~30–40 rows (all valid students).
    - DOBs in `MM/DD/YYYY` (e.g., `06/29/2005`).
    - Genders (`Male`, `Female`).
    - Correct defaults (e.g., `Provider ID=2`).
- **Console Output**:
  - Check `Source DataFrame Columns` for suffixed names (e.g., `First Name.1`).
  - Verify `Processing [Student] Student` logs for valid `First Name`, `Last Name`, `DOB`, `Gender`.
  - Confirm `Added row for...` messages for each student.
  - Note DOB parsing warnings (e.g., `Warning: Failed to parse DOB...`).

## Troubleshooting
- **Empty Output File**:
  - **Cause**: Mismatched column names (e.g., expecting `First Name.1` but source differs).
  - **Solution**: Check `Source DataFrame Columns` log. Update `student_sets` in `transform_excel_data.py` to match.
- **Missing Students**:
  - **Cause**: Second, Third, or Fourth students skipped due to empty `First Name`/`Last Name`.
  - **Solution**: Verify source file data. Ensure `Processing` logs show correct names.
- **Empty DOB or Gender**:
  - **Cause**: Incorrect `dob_key`/`gender_key` (e.g., `Second Student Birth Date` missing).
  - **Solution**: Confirm `dob_key` uses `First Student Birth Date`, etc. Check parsing warnings.
- **Duplicate Column Issues**:
  - **Cause**: Pandas appends `.1`, `.2`, `.3` to duplicate `First Name`, etc.
  - **Solution**: Ensure `transform_excel_data.py` uses suffixed names. Verify source file structure.
- **Script Errors**:
  - Share console error messages or stack traces.
- **Unexpected Output**:
  - Compare output rows against expected students.
  - Share console logs and problematic rows.

## Contributing
Contributions are welcome! To contribute:
1. Fork the repository.
2. Create a feature branch (`git checkout -b feature/your-feature`).
3. Commit changes (`git commit -m "Add your feature"`).
4. Push to the branch (`git push origin feature/your-feature`).
5. Open a Pull Request.

Please ensure code follows Python best practices and includes tests or documentation updates.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contact
For questions or support, contact the USCG Auxiliary Flotilla 65 team or open an issue on the [GitHub repository](https://github.com/jjborie/Flotilla65FWCBoatingEdTransformer).