# Data Cleaning & Reporting Automation

## Objective
Automate a data-cleaning and reporting workflow using Python and Pandas.

## Key Features
- Standardizes inconsistent text such as city, category and payment status.
- Converts currency and numeric columns into usable numeric values.
- Detects invalid dates and missing values.
- Removes duplicate Order IDs.
- Fills missing numeric/date values using median values.
- Calculates Revenue = Quantity × Unit Price.
- Generates summary reports by city, category, product, payment status and date.
- Creates visual summaries and an Excel dashboard.

## Dataset
The included `data/raw_sales_data.csv` is a deliberately messy sample sales dataset containing:
- Missing quantity values
- Duplicate order records
- Currency symbols
- Inconsistent capitalization and spaces
- An invalid date
- Mixed date formats

## Output
After running `python automate_report.py`:
- `data/cleaned_sales_data.csv`
- `reports/automated_data_cleaning_report.xlsx`
- Charts inside `charts/`

## Tools Used
Python, Pandas, Matplotlib, OpenPyXL and Excel.

## Sample Results
- Original rows: 50
- Duplicate Order IDs removed: 0
- Missing values handled: 5
- Final clean rows: 50
- Total revenue: ₹1,239,100.00
- Average order value: ₹24,782.00

## Submission Note
This project demonstrates an end-to-end automated data preprocessing and reporting workflow and can be extended to real CSV/Excel files by replacing the input dataset.
