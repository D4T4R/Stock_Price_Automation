from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import yfinance as yf
import openpyxl
from pathlib import Path
#from nsetools import Nse

def fetch_yfinance(stock_symbols):
    """
    Fetches current stock prices using Yahoo Finance and rounds them to 2 decimal places.
   
    """
    prices = {}
    for symbol in stock_symbols:
        try:
            stock = yf.Ticker(symbol)
            price = stock.history(period="1d")['Close'].iloc[-1]
            prices[symbol] = round(price, 2)
        except Exception as e:
            print(f"Error fetching data for {symbol}: {e}")
    return prices

def update_excel(file_path, column, start_row, stock_name_to_scrip, prices):
    """
    Updates the specified column in an Excel file with the latest prices and applies color formatting to column K.
    
    """
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active

        for i, row in enumerate(sheet.iter_rows(min_row=start_row), start=start_row):
            stock_name = row[0].value  # Col. A
            if stock_name in stock_name_to_scrip:
                symbol = stock_name_to_scrip[stock_name]
                if symbol in prices:
                    sheet[f"{column}{i}"] = prices[symbol]

        column_k_index = openpyxl.utils.column_index_from_string("K")
        for row in sheet.iter_rows(min_row=start_row, min_col=column_k_index, max_col=column_k_index):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    if cell.value > 100:
                        cell.font = Font(color="00FF00")  # Colour 1
                    elif cell.value < 0:
                        cell.font = Font(color="FF0000")  # Colour 2

        workbook.save(file_path)
        print(f"Successfully updated {file_path}")
    except Exception as e:
        print(f"Error updating Excel file: {e}")

if __name__ == "__main__":
    file_path = Path("/blah-blah/abcd-xyz/Final_P&L_auto.xlsx")  # Actual file path
    column_to_update = "F"
    start_row = 2
    stock_name_to_scrip = {
        
        "ASIAN PAINTS": "ASIANPAINT.NS",
        "BRITANNIA INDUSTRIES": "BRITANNIA.NS",
        "HAPPIEST MINDS TECH": "HAPPSTMNDS.NS",
        "HCL TECHNOLOGIES": "HCLTECH.NS",
        "ITC": "ITC.NS",
        "MAHINDRA & MAHINDRA": "M&M.NS",
        "PTC INDIA": "PTC.NS",
        "TATA CHEMICALS": "TATACHEM.NS",
        "TATA ELXSI": "TATAELXSI.NS",
        "TATA POWER": "TATAPOWER.NS",
        "TATA STEEL": "TATASTEEL.NS",
        "INFOSYS": "INFY.NS",
        "WIPRO": "WIPRO.NS",
        "ADANI PORTS": "ADANIPORTS.NS",
        "DELTACORP": "DELTACORP.NS",
        "DRREDDY": "DRREDDY.NS",
        "GRASIM": "GRASIM.NS",
        "HAVELLS": "HAVELLS.NS",
        "INDIAN HOTELS": "INDHOTEL.NS",
        "SIEMENS": "SIEMENS.NS",
        "IRCTC": "IRCTC.NS",
        "STATE BANK OF INDIA": "SBIN.NS",
        "TRENT": "TRENT.NS",
        "LIC INDIA": "LICI.NS",
        "TCS": "TCS.NS"
    }

    # Name extraction
    stock_names = []
    if file_path.exists():
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        stock_names = [sheet[f"A{row}"].value for row in range(start_row, sheet.max_row + 1)]

    # Fetch
    scrips_to_fetch = [stock_name_to_scrip[name] for name in stock_names if name in stock_name_to_scrip]
    print("Fetching live data from Yahoo Finance...")
    stock_prices = fetch_yfinance(scrips_to_fetch)
    print("Fetched Prices:", stock_prices)

    # Update Excel file
    update_excel(file_path, column_to_update, start_row, stock_name_to_scrip, stock_prices)