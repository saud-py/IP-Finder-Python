import pandas as pd
import requests
import openpyxl
from openpyxl.styles import Border, Side
import os
import sys
import logging
import time
from datetime import datetime
from requests.exceptions import RequestException

def get_ip_details(ip_address, max_retries=3, timeout=5):
    """Fetch IP details using ip-api.com (free API) with retry logic"""
    retries = 0
    while retries < max_retries:
        try:
            logging.debug(f"Fetching details for IP: {ip_address} (Attempt {retries+1}/{max_retries})")
            response = requests.get(f"http://ip-api.com/json/{ip_address}", timeout=timeout)
            data = response.json()
            
            if data["status"] == "success":
                result = {
                    "IP Address": ip_address,
                    "Country": data.get("country", "Unknown"),
                    "Region/State": data.get("regionName", "Unknown"),
                    "City": data.get("city", "Unknown"),
                    "ISP/Organization": data.get("isp", "Unknown"),
                    "Timezone": data.get("timezone", "Unknown")
                }
                logging.debug(f"Successfully retrieved details for {ip_address}")
                return result
            else:
                logging.warning(f"Failed to retrieve details for IP {ip_address}: {data.get('message', 'Unknown error')}")
                return {
                    "IP Address": ip_address,
                    "Country": "Failed to retrieve",
                    "Region/State": "Failed to retrieve",
                    "City": "Failed to retrieve",
                    "ISP/Organization": "Failed to retrieve",
                    "Timezone": "Failed to retrieve"
                }
        except (RequestException, TimeoutError) as e:
            retries += 1
            error_msg = f"Error fetching details for IP {ip_address} (Attempt {retries}/{max_retries}): {str(e)}"
            logging.warning(error_msg)
            print(error_msg)
            
            if retries < max_retries:
                wait_time = 2 * retries  # Exponential backoff
                logging.info(f"Waiting {wait_time} seconds before retry...")
                time.sleep(wait_time)
            else:
                logging.error(f"Max retries reached for IP {ip_address}")
                return {
                    "IP Address": ip_address,
                    "Country": "Connection Error",
                    "Region/State": "Connection Error",
                    "City": "Connection Error",
                    "ISP/Organization": "Connection Error",
                    "Timezone": "Connection Error"
                }
        except Exception as e:
            error_msg = f"Unexpected error fetching details for IP {ip_address}: {str(e)}"
            logging.error(error_msg)
            print(error_msg)
            return {
                "IP Address": ip_address,
                "Country": "Error",
                "Region/State": "Error",
                "City": "Error",
                "ISP/Organization": "Error",
                "Timezone": "Error"
            }

def apply_border_to_cells(worksheet):
    """Apply borders to all cells with data"""
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = thin_border

def setup_logging():
    """Set up logging to file and console"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = f"ip_finder_{timestamp}.log"
    
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    
    # Add specific logger for requests and urllib3 to reduce noise
    logging.getLogger("requests").setLevel(logging.WARNING)
    logging.getLogger("urllib3").setLevel(logging.WARNING)
    
    logging.info(f"Log file created: {log_file}")
    return log_file

def main():
    # Set up logging
    log_file = setup_logging()
    
    # Check if input file is provided
    if len(sys.argv) < 2:
        logging.error("No input file provided")
        print("Usage: python ip_finder.py <input_excel_file>")
        print("Example: python ip_finder.py input.xlsx")
        return
    
    input_file = sys.argv[1]
    logging.info(f"Input file: {input_file}")
    
    # Check if file exists
    if not os.path.exists(input_file):
        logging.error(f"File '{input_file}' not found")
        print(f"Error: File '{input_file}' not found.")
        return
    
    try:
        # Read the Excel file
        logging.info("Reading Excel file...")
        df = pd.read_excel(input_file)
        
        # Use the specified column name "dstAddr" for IP addresses
        ip_column = "dstAddr"
        
        if ip_column not in df.columns:
            logging.error(f"Column '{ip_column}' not found in the Excel file")
            logging.info(f"Available columns: {df.columns.tolist()}")
            print(f"Error: Column '{ip_column}' not found in the Excel file.")
            print("Available columns:", df.columns.tolist())
            return
        
        logging.info(f"Using IP column: {ip_column}")
        logging.info(f"Processing {len(df)} IP addresses...")
        print(f"Using IP column: {ip_column}")
        print(f"Processing {len(df)} IP addresses...")
        
        # Create a list to store results
        results = []
        
        # Process each IP address
        for index, row in df.iterrows():
            ip = str(row[ip_column]).strip()
            if ip and ip != "nan":
                logging.info(f"Processing IP {index+1}/{len(df)}: {ip}")
                print(f"Processing IP {index+1}/{len(df)}: {ip}")
                
                # Add delay between requests to avoid rate limiting (ip-api.com allows 45 requests per minute)
                if index > 0 and index % 40 == 0:
                    wait_time = 60
                    logging.info(f"Pausing for {wait_time} seconds to avoid rate limiting...")
                    print(f"Pausing for {wait_time} seconds to avoid rate limiting...")
                    time.sleep(wait_time)
                
                try:
                    ip_details = get_ip_details(ip)
                    results.append(ip_details)
                    logging.info(f"Retrieved details for {ip}: {ip_details['Country']}, {ip_details['City']}")
                except KeyboardInterrupt:
                    logging.warning("Process interrupted by user. Saving results collected so far...")
                    print("Process interrupted by user. Saving results collected so far...")
                    break
            else:
                logging.warning(f"Skipping empty IP at row {index+1}")
                print(f"Skipping empty IP at row {index+1}")
        
        # Create a DataFrame from results
        results_df = pd.DataFrame(results)
        
        # Generate output filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"ip_details_{timestamp}.xlsx"
        
        # Save to Excel
        logging.info(f"Saving results to {output_file}")
        results_df.to_excel(output_file, index=False)
        
        # Apply borders
        logging.info("Applying borders to cells")
        workbook = openpyxl.load_workbook(output_file)
        worksheet = workbook.active
        apply_border_to_cells(worksheet)
        workbook.save(output_file)
        
        logging.info(f"Results saved to {output_file}")
        print(f"Results saved to {output_file}")
        logging.info(f"Log file available at: {log_file}")
        print(f"Log file available at: {log_file}")
        
    except Exception as e:
        error_msg = f"An error occurred: {str(e)}"
        logging.error(error_msg)
        print(error_msg)

if __name__ == "__main__":
    main()
