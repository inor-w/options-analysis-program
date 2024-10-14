import math
import pandas as pd
import logging
import sys
from datetime import datetime
from tkinter import messagebox
from scipy.stats import norm
from utils import clean_date_string

def calc_variables(data_sheet):
    """
    Calculate required variables from the data sheet for further processing.
    """
    variables = {key: None for key in [
        "LastStockPrice", "StrikePrice", "IVMid", "Expiration", "BidOptionPrice", "Volume",
        "MyOptionQuantity", "PurchasedPrice", "PercentGainLoss", "Description", "EarningsDate",
        "Delta", "Gamma", "Vega", "Theta", "Rho", "OpenInt", "Account", "Type", "TradeTime"
    ]}
    try:
        # Check if data_sheet is initialized
        if data_sheet is None:
            raise Exception("Data sheet is not initialized")

        # Extract Account information from cell A2
        variables["Account"] = data_sheet['A2'].value

        # Clean and parse TradeTime from cell A1
        trade_time_str = clean_date_string(data_sheet['A1'].value)
        logging.info(f"TradeTime string to parse: '{trade_time_str}'")

        # Parse TradeTime into a time object
        try:
            date_time_obj = datetime.strptime(trade_time_str, "%m/%d/%Y %I:%M:%S %p")
            variables["TradeTime"] = date_time_obj.time()
            logging.info(f"SUCCESS: The TradeTime is {variables['TradeTime']}")
        except (ValueError, TypeError):
            logging.error("TradeTime is invalid or not in the correct format")
            variables["TradeTime"] = None

        # Define market open and close times
        market_open_time = datetime.strptime("9:30 AM", "%I:%M %p").time()
        market_close_time = datetime.strptime("4:00 PM", "%I:%M %p").time()

        # Mapping from keywords in the data sheet to variable names
        keyword_to_variable = {
            "Last": "LastStockPrice", "StrikePrice": "StrikePrice", "IVMid": "IVMid", "Expiration": "Expiration",
            "Bid": "BidOptionPrice", "OpenInt": "OpenInt", "Volume": "Volume", "Quantity": "MyOptionQuantity",
            "PurchasePrice": "PurchasedPrice", "PercentGainLoss": "PercentGainLoss", "Description": "Description",
            "EarningsDate": "EarningsDate", "Delta": "Delta", "Gamma": "Gamma", "Vega": "Vega",
            "Theta": "Theta", "Rho": "Rho", "Strategies": "Type"
        }

        # Extract variables from data_sheet based on keywords
        custom_target_rows = {"Last": 1}  # For 'Last', the value is in the next row
        for row_index, row in enumerate(data_sheet.iter_rows(values_only=True), start=1):
            for col_index, cell_value in enumerate(row, start=1):
                if cell_value in keyword_to_variable:
                    # Determine the target row to get the value
                    target_row = row_index + custom_target_rows.get(cell_value, 2)
                    if target_row <= data_sheet.max_row:
                        value = data_sheet.cell(row=target_row, column=col_index).value
                        # If the keyword is 'Expiration', parse the date
                        if cell_value == "Expiration" and value:
                            try:
                                value = pd.to_datetime(value)
                            except ValueError:
                                logging.warning(f"Could not parse expiration date: {value}")
                        variables[keyword_to_variable[cell_value]] = value

        # Validate required variables
        required_vars = ["Expiration", "LastStockPrice", "StrikePrice", "IVMid"]
        for var in required_vars:
            if variables[var] is None:
                logging.error(f"Required variable '{var}' is missing")
                messagebox.showerror("ERROR", f"Required variable '{var}' is missing in the data")
                sys.exit()

        # Calculate additional variables
        today = datetime.now()
        expiration_date = variables["Expiration"]
        days_until_expiration = (expiration_date - today).days + 1  # Days left until expiration
        formatted_expiration = expiration_date.strftime("%m/%d/%Y")  # Format expiration date

        # Parse dateRetrieved (same as TradeTime)
        dateRetrieved_str = clean_date_string(data_sheet['A1'].value)
        logging.info(f"dateRetrieved string to parse: '{dateRetrieved_str}'")
        dateRetrieved = pd.to_datetime(dateRetrieved_str, format='%m/%d/%Y %I:%M:%S %p')
        formatted_dateRetrieved = dateRetrieved.strftime("%m/%d/%Y")

        # Determine tradeTime status (during market hours or not)
        if variables["TradeTime"] and market_open_time <= variables["TradeTime"] <= market_close_time:
            tradeTime = variables["TradeTime"].strftime("%I:%M:%S %p")
        else:
            tradeTime = "Market Closed"

        # Calculate ITM and OTM probabilities
        itm_probability, otm_probability = calculate_probs(
            variables["LastStockPrice"],
            variables["StrikePrice"],
            variables["IVMid"],
            days_until_expiration
        )

        # Return variables and calculation results
        calculation_results = {
            'variables': variables,
            'formatted_expiration': formatted_expiration,
            'days_until_expiration': days_until_expiration,
            'formatted_dateRetrieved': formatted_dateRetrieved,
            'tradeTime': tradeTime,
            'itm_probability': itm_probability,
            'otm_probability': otm_probability
        }

        return calculation_results

    except Exception as e:
        logging.error(f"An unexpected error occurred: {str(e)}")
        messagebox.showerror("ERROR", f"An unexpected error occurred: {str(e)}")
        sys.exit()

def calculate_probs(last_stock_price, strike_price, iv_mid, days_until_expiration):
    """
    Calculate the in-the-money (ITM) and out-of-the-money (OTM) probabilities.
    """
    try:
        # Convert inputs to appropriate numeric types
        S = float(last_stock_price)  # Current stock price
        K = float(strike_price)      # Strike price
        r = 0.039  # Risk-free interest rate
        sigma = float(iv_mid.strip('%')) / 100  # Convert IV from percentage to decimal
        T = days_until_expiration / 365  # Time to expiration in years

        # Calculate d2 using Black-Scholes formula components
        d2 = (math.log(S / K) + (r - 0.5 * sigma ** 2) * T) / (sigma * math.sqrt(T))

        # Calculate probabilities using cumulative distribution function
        otm_probability = norm.cdf(-d2)  # Probability of being out-of-the-money
        itm_probability = 1 - otm_probability  # Probability of being in-the-money

        logging.info(f"Calculated ITM probability: {itm_probability:.2%}")
        logging.info(f"Calculated OTM probability: {otm_probability:.2%}")

        return itm_probability, otm_probability
    except Exception as e:
        logging.error(f"Error calculating probabilities: {str(e)}")
        messagebox.showerror("ERROR", f"Error calculating probabilities: {str(e)}")
        sys.exit()
