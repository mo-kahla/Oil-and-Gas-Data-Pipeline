# Oil and Gas Data Pipeline

**Note: The data and company names referenced in this project have been altered or concealed for privacy and security reasons.**

This project is a data pipeline designed to process and consolidate production data for an oil and gas company. The pipeline extracts, cleans, transforms, and loads data from various Excel and CSV files into a SQLite database.

## Project Structure

The project consists of the following main scripts:

### 1. `my_update.py`
This script processes the "Remarks" sheet from the Excel files to extract and transform ESP sensor data.

- **Identify ESP Sensor Data**: Uses regular expressions to identify cells containing "ESP sensor data" and extracts relevant rows and columns.
- **Rename Columns**: Renames columns to standardized names and converts data to numeric types where applicable.
- **Save Processed Data**: Saves the processed data for further analysis.

### 2. `try.py`
This script processes well injection data from the Excel files.

- **Identify Well and Total Rows**: Uses regular expressions to identify cells containing "WELL NAME" and "TOTAL".
- **Extract and Transform Data**: Extracts data between identified rows, cleans and standardizes well names and zones, and calculates unique IDs for wells.
- **Save Processed Data**: Saves the processed data to a CSV file named `inj.csv`.

### 3. `dpr.py`
This script handles the extraction and processing of production data from Excel files.

- **Unhide Sheets**: Ensures that the specified sheet in the Excel files is visible.
- **Read Data**: Reads data from the "PRODUCTION DPR Report" sheet in the Excel files.
- **Transform Data**: Cleans and transforms the extracted data by identifying specific cells and columns, renaming them, and handling missing values.

### 4. `control.py`
This script orchestrates the execution of various update scripts.

- **Execute Scripts**: Uses the subprocess module to run multiple scripts in sequence:
  - `ofm_update_pressure.py`
  - `ofm_update_production.py`
  - `ofm_update_test.py`
  - `ofm_update_events.py`
- **Print Status**: Prints a confirmation message after each script completes.

### 5. `sql_update.py`
This script loads various CSV files into a SQLite database.

- **Read and Rename Columns**: Reads CSV files and renames columns for consistency.
- **Load Data into SQLite**: Loads data into appropriate tables in the SQLite database.
- **Close Connection**: Closes the database connection after data is loaded.

## Key Achievements

- **Efficiency Boost**: Developed a robust data migration pipeline utilizing Python, SQL, and regex, seamlessly transferring data to the database.
- **Automation**: Automated the ETL process across thousands of rows of data, reducing the manual workload by 30%.
- **Time Savings**: Significantly reduced manual workload by 1.5 hours per day through efficient data transformation and loading processes.

## Usage 
A batch file (run_pipeline.bat) is provided to automate this process

## License

This project is licensed under the MIT License. See the LICENSE file for details.
