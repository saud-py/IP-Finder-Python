# IP Finder

A Python script that reads IP addresses from an Excel file, fetches their details, and outputs the results to a formatted Excel file.

## Features

- Reads IP addresses from an Excel file
- Fetches the following details for each IP:
  - Country
  - Region/State
  - City
  - ISP/Organization
  - Timezone
- Outputs results to a new Excel file with bordered cells

## Requirements

- Python 3.6+
- Required packages: pandas, requests, openpyxl

## Installation

1. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

## Usage

Run the script with the input Excel file as an argument:

```
python ip_finder.py input.xlsx
```

Where `input.xlsx` is your Excel file containing IP addresses.

The script will:
1. Automatically detect the column containing IP addresses
2. Fetch details for each IP
3. Create a new Excel file with the results (named `ip_details_YYYYMMDD_HHMMSS.xlsx`)

## Notes

- The script uses the free IP-API.com service which has rate limits (45 requests per minute)
- For large datasets, the script may take some time to complete