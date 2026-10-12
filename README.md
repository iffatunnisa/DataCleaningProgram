# Data Cleaning — Balmon Aceh

A web-based data cleaning application built with **Python, Pandas, OpenPyXL, and Gradio** to simplify the preprocessing of telecommunications measurement data for further analysis and database integration.

The application provides two main features: Speed Testing data cleaning and Compile Summary header standardization.

## Features

### 1. Speed Testing Data Cleaning

Processes raw Excel files containing Speed Testing measurements and exports the cleaned data as a CSV file.

**Processing steps:**
- Reads the `Speed_Testing` worksheet with the header located on the second row.
- Selects the required columns:
  - `Home Operator`
  - `Application Throughput DL`
  - `Application Throughput UL`
  - `Latitude`
  - `Longitude`
- Converts downlink and uplink throughput values into numeric data, including values that use decimal commas.
- Converts latitude and longitude into numeric values.
- Removes rows with missing or invalid coordinates.
- Preserves other rows, including rows with missing throughput values.
- Exports the result as `SpeedTesting_Cleaned.csv`.

**Supported input formats:** `.xls`, `.xlsx`, `.xlsm`

### 2. Compile Summary Header Standardization

Simplifies multi-row Excel headers into a single row of concise, database-friendly column names.

**Processing steps:**
- Reads the `Compile_Summary` worksheet when available; otherwise, uses the active worksheet.
- Combines the first and second header rows.
- Abbreviates recognized measurement categories, including FTP, HTTP, Speed Test, browsing, voice calls, and signal quality.
- Standardizes common metric names and measurement units.
- Generates unique column names by appending numeric suffixes to duplicates.
- Preserves the underlying data rows without filtering them.
- Retains the header styling and adjusts merged cells in the worksheet.
- Exports the result as `Compile_Summary_Cleaned.xlsx`.

This feature is designed to make column names easier to manage when preparing data for PostgreSQL or other analytical workflows.

## Tech Stack

- **Python** — application logic
- **Gradio** — interactive web interface
- **Pandas** — tabular data processing
- **OpenPyXL** — Excel workbook and header processing

## Project Structure

```text
DataCleaningProgram/
├── app.py
├── requirement.txt
├── .gitignore
└── README.md
```

## Installation

### Prerequisites

- Python 3.10 or later recommended
- pip

### 1. Clone the repository

```bash
git clone https://github.com/iffatunnisa/DataCleaningProgram.git
cd DataCleaningProgram
```

### 2. Create a virtual environment (recommended)

**Windows:**

```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. Install dependencies

The dependency file in this repository is named `requirement.txt`.

```bash
pip install -r requirement.txt
```

### 4. Run the application

```bash
python app.py
```

Open the local URL displayed in the terminal, typically:

```text
http://127.0.0.1:7860
```

## How to Use

### Speed Testing

1. Open the **Speed Testing** tab.
2. Upload a raw Excel file containing a `Speed_Testing` worksheet.
3. Click **Proses Speed Testing**.
4. Review the processing status and row counts.
5. Download the resulting `SpeedTesting_Cleaned.csv` file.

### Compile Summary

1. Open the **Compile Summary** tab.
2. Upload an Excel file containing two header rows.
3. Click **Proses Compile Summary**.
4. Review the processing status.
5. Download the resulting `Compile_Summary_Cleaned.xlsx` file.

## Output Files

| Feature | Output file | Format |
|---|---|---|
| Speed Testing | `SpeedTesting_Cleaned.csv` | CSV |
| Compile Summary | `Compile_Summary_Cleaned.xlsx` | Excel |

Output files are generated when the corresponding processing operation is executed.

## Notes and Limitations

- Input files must follow the expected worksheet and header structure.
- Speed Testing cleaning removes rows with invalid coordinates but does not remove rows solely because throughput values are missing.
- Compile Summary standardizes headers; it does not perform data filtering or comprehensive data validation.
- The application currently processes one uploaded file at a time per operation.
- The application runs locally by default. Network access, deployment, and shared-server configuration require additional setup.
- Avoid uploading confidential or sensitive operational data to public hosting services or repositories.

## License

No license has been specified yet. All rights to the source code and related materials remain with their respective owners unless otherwise stated.
