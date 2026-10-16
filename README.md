# Data Cleaning & Reporting Automation

A beginner-friendly Python project that automatically cleans messy sales data and generates an Excel report and charts.

## Features
- Removes exact duplicate rows
- Trims extra spaces and standardizes region names
- Converts date, quantity, and price columns to suitable data types
- Corrects missing/invalid quantities and missing values
- Calculates sales for each order
- Exports cleaned data as CSV
- Creates an Excel workbook with data-quality metrics, region summary, product summary, daily sales, and cleaned data
- Saves sales-by-region and sales-by-product charts

## Project structure
```text
data_cleaning_reporting_automation/
├── data/
│   └── sales_data_raw.csv
├── outputs/                  # Generated after running the script
├── reports/                  # Generated charts saved here
├── main.py
├── requirements.txt
└── README.md
```

## Requirements
Python 3.10 or newer is recommended.

## Run the project
1. Install Python.
2. Open a terminal in this project folder.
3. Install the libraries:
   ```bash
   python -m pip install -r requirements.txt
   ```
4. Run:
   ```bash
   python main.py
   ```
5. Open `outputs/automated_sales_report.xlsx` and `outputs/cleaned_sales_data.csv`.
   The chart images will be in `reports/`.

## How the cleaning works
1. Loads the raw CSV file with pandas.
2. Removes rows that are exact duplicates.
3. Strips extra spaces and normalizes region labels.
4. Parses dates and numeric columns; invalid values are treated as missing.
5. Replaces missing numeric values with the column median, missing dates with the most common valid date, and missing customer/email/salesperson values with `Not Provided`.
6. Calculates `Sales = Quantity × Unit_Price`.
7. Produces summaries and visual reports.

## Customizing with your own data
Replace `data/sales_data_raw.csv` with your dataset. Keep the same column names (`Order_ID`, `Date`, `Customer`, `Region`, `Product`, `Quantity`, `Unit_Price`, `Salesperson`, `Email`) or update the column names in `main.py`.

## Notes
- This is a demonstration dataset, not real customer information.
- Median imputation is a simple baseline. For a real business dataset, review the missing-data strategy with a domain expert before using reports for decisions.
