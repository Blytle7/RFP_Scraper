import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
from datetime import datetime


# Function to load sites from Excel
def load_sites_from_excel(excel_file, sheet_name):
    """Load the list of sites from a specific sheet in the Excel file."""
    # Load the Excel sheet into a pandas DataFrame
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=0)

    # Print the actual column names from the sheet to debug
    print("Columns in the sheet:", df.columns)  # This will print the column names

    # Get the Description and Link columns and convert to a list
    sites = df[['Description', 'Link']].dropna().values.tolist()  # Adjusted for "Description" and "Link" headers
    return sites


# Function to scrape a website for RFQs
def scrape_rfq(website_url):
    """Scrape a website for any new RFQs. Returns True if new RFQs are found."""
    try:
        # Send a GET request to the website
        response = requests.get(website_url, timeout=10)
        response.raise_for_status()  # Raise an exception for 4xx/5xx errors

        # Use BeautifulSoup to parse the page content
        soup = BeautifulSoup(response.text, 'html.parser')

        # Check if the page contains any RFQs (simple keyword search)
        keywords = ['RFQ', 'Request for Quote', 'Bid']
        for keyword in keywords:
            if keyword.lower() in soup.get_text().lower():
                return True
        return False
    except requests.exceptions.RequestException as e:
        print(f"Error fetching {website_url}: {e}")
        return False


# Function to check all sites for new RFQs and save to an Excel file
def check_for_rfq(excel_file, sheet_name, output_file):
    """Check all sites for new RFQs and save results to a new Excel file."""
    sites = load_sites_from_excel(excel_file, sheet_name)

    # List to store results
    results = []

    # Loop through each site in the list
    for description, link in sites:
        print(f"Checking {description} ({link}) for new RFQs...")

        # Scrape the site and check for RFQs
        rfq_found = scrape_rfq(link)

        # If RFQs are found, save the result
        if rfq_found:
            results.append([description, link, "New RFQ Found", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
        else:
            results.append([description, link, "No New RFQs", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])

        # Add a delay between requests to avoid being blocked
        time.sleep(2)

    # Save the results to an Excel file
    results_df = pd.DataFrame(results, columns=["Description", "Link", "Status", "Checked At"])
    results_df.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")


# Main function to initiate the RFQ check
def main():
    # Input file and sheet name
    excel_file = 'RFQ Hunt.xlsx'
    sheet_name = 'Sites_Test'  # The sheet name containing the site info
    output_file = 'RFQ_Results.xlsx'  # File to save the results

    # Check for new RFQs and save the results
    check_for_rfq(excel_file, sheet_name, output_file)


# Run the main function
if __name__ == "__main__":
    main()
