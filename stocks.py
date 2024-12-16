from nsetools import Nse
import openpyxl
from pathlib import Path

def fetch_nse_prices(scrips):
    """
    Fetches current market prices for a list of scrip symbols from NSE.
    
    """
    nse = Nse()
    prices = {}
    for scrip in scrips:
        try:
            data = nse.get_quote(scrip)
            if data:
                prices[scrip] = data['lastPrice']
        except Exception as e:
            print(f"Error fetching data for {scrip}: {e}")
    return prices

def update_excel(file_path, column, start_row, stock_name_to_scrip, prices):
    """
    Updates the specified column in an Excel file with the latest prices.
    
    Args:
        file_path (str): Path to the Excel file.
        column (str): Column letter to update (e.g., 'F').
        start_row (int): Row number to start updating from (e.g., 2).
        stock_name_to_scrip (dict): Mapping of stock names to NSE scrip symbols.
        prices (dict): Latest prices fetched from NSE.
    """
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active

        for i, row in enumerate(sheet.iter_rows(min_row=start_row), start=start_row):
            stock_name = row[0].value  # Target -> Col A
            if stock_name in stock_name_to_scrip:
                scrip = stock_name_to_scrip[stock_name]
                if scrip in prices:
                    sheet[f"{column}{i}"] = prices[scrip]

        workbook.save(file_path)
        print(f"Successfully updated {file_path}")
    except Exception as e:
        print(f"Error updating Excel file: {e}")

if __name__ == "__main__":
    file_path = Path("/blahblah/abcdxyz/Final_P&L_auto.xlsx")  # Replace with file path
    column_to_update = "F"
    start_row = 2

    # NSE scrip symbols
    stock_name_to_scrip = {
        "ASIAN PAINTS": "ASIANPAINT",
        "BRITANNIA INDUSTRIES": "BRITANNIA",
        "HAPPIEST MINDS TECH": "HAPPSTMNDS",
        "HCL TECHNOLOGIES": "HCLTECH",
        "ITC": "ITC",
        "MAHINDRA & MAHINDRA": "M&M",
        "PTC INDIA": "PTC",
        "TATA CHEMICALS": "TATACHEM",
        "TATA ELXSI": "TATAELXSI",
        "TATA POWER": "TATAPOWER",
        "TATA STEEL": "TATASTEEL",
        "INFOSYS": "INFY",
        "WIPRO": "WIPRO",
        "ADANI PORTS": "ADANIPORTS",
        "DELTACORP": "DELTACORP",
        "DRREDDY": "DRREDDY",
        "GRASIM": "GRASIM",
        "HAVELLS": "HAVELLS",
        "INDIAN HOTELS": "INDHOTEL",
        "SIEMENS": "SIEMENS",
        "IRCTC": "IRCTC",
        "STATE BANK OF INDIA": "SBIN",
        "TRENT": "TRENT",
        "LIC INDIA": "LICI",
        "TCS": "TCS"
        #"RelianceIndustries": "RELIANCE",
        #"HDFC Bank": "HDFCBANK"
    }

    # Name extraction
    stock_names = []
    if file_path.exists():
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        stock_names = [sheet[f"A{row}"].value for row in range(start_row, sheet.max_row + 1)]

    # Fetch
    scrips_to_fetch = [stock_name_to_scrip[name] for name in stock_names if name in stock_name_to_scrip]
    print("Fetching live data from NSE...")
    stock_prices = fetch_nse_prices(scrips_to_fetch)
    print("Fetched Prices:", stock_prices)

    # Update Excel file 
    update_excel(file_path, column_to_update, start_row, stock_name_to_scrip, stock_prices)
