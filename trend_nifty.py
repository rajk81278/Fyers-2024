import credentials as crs
import pandas as pd
import pandas_ta as ta
from fyers_api import fyersModel
import openpyxl
import datetime as dt
import time
import yfinance as yf

import xlwings as xw


from fyers_api import fyersModel


# Import the required module from the fyers_apiv3 package
from fyers_apiv3 import fyersModel
import webbrowser
import time

# Replace these values with your actual API credentials
client_id = crs.client_id
secret_key = crs.secret_key
redirect_uri = crs.redirect_uri
response_type = "code"  
state = "sample_state"

# Create a session model with the provided credentials
session = fyersModel.SessionModel(
    client_id=client_id,
    secret_key=secret_key,
    redirect_uri=redirect_uri,
    response_type=response_type

)

# Generate the auth code using the session model
response = session.generate_authcode()

# Print the auth code received in the response
print(response)

webbrowser.open(response,new=1)


newurl = input("Enter the url: ")
auth_code = newurl[newurl.index('auth_code=')+10:newurl.index('&state')]
print(auth_code)


grant_type = "authorization_code" 
session = fyersModel.SessionModel(
    client_id=client_id,
    secret_key=secret_key, 
    redirect_uri=redirect_uri, 
    response_type=response_type, 
    grant_type=grant_type
)

# Set the authorization code in the session object
session.set_token(auth_code)

# Generate the access token using the authorization code
response = session.generate_token()

# Print the response, which should contain the access token and other details
print(response)


# There can be two cases over here you can successfully get the acccessToken over the request or you might get some error over here. so to avoid that have this in try except block
try: 
    access_token = response["access_token"]
    with open('access.txt','w') as k:
        k.write(access_token)
except Exception as e:
    print(e,response)  ## This will help you in debugging then and there itself like what was the error and also you would be able to see the value you got in response variable. instead of getting key_error for unsuccessfull response.






with open('access.txt') as f:
    # Read the contents of the file into a variable
    access_token = f.read()
    # Print the names
    # print(access_token)


def process_and_export_to_excel():
    wb = xw.Book('AT_Glance_Nifty.xlsx')



    sheet = wb.sheets['Sheet1']
    sheet.range('A1:Z100').color = (0, 0, 0)  # Black background
    sheet.range('A1:Z100').api.Font.Color = 0xFFFFFF  # White font

    # sheet.range('A2:E2').color = (0, 26, 51)
    # sheet.range('G2:J2').color = (173, 216, 230)
    sheet.range('A2:E2').api.Font.Bold = True
    sheet.range('G2:J2').api.Font.Bold = True
    sheet['A2'].value = 'Name'
    sheet['B2'].value = 'Price'
    sheet['C2'].value = 'Value'
    sheet['D2'].value = 'Percentage'
    sheet['E2'].value = 'Remarks'


    sheet['A3'].value = 'Nifty Spot'
    sheet['A4'].value = 'Nifty Future'
    sheet['A5'].value = 'Spread'
    sheet['A11'].value = 'AVWAP'
    sheet['A12'].value = 'VWAP'
    sheet['A13'].value = 'RSI'
    sheet['A14'].value = 'ATR'
    sheet['A15'].value = 'Week High / Low'
    sheet['A16'].value = 'Monthly High / Low'
    sheet['A17'].value = '52 Week High / Low'

    sheet['G2'].value = 'Strike'
    sheet['H2'].value = 'CE Price'
    sheet['I2'].value = 'PE Price'
    sheet['J2'].value = 'Total Premium'


    def fetch_nifty_spot():
        # Fetching spot and future data using Fyers quotes  NSE:NIFTY24OCTFUT
        data = {
            "symbols":  'NSE:NIFTY50-INDEX'
        }
        response = fyers.quotes(data=data)
        
        spot_data = response  # Nifty Spot data
        # future_data = response['d'][1]  # Nifty Future data
        # print(spot_data)
        # print(future_data)
    
        return spot_data # future_data


    fyers = fyersModel.FyersModel(client_id= 'XZ5S16H9QE-100', token=access_token,is_async=False, log_path="")

    # Make a request to get the funds information
    response = fyers.funds()



    # Function to fetch live Nifty spot and future prices
    def fetch_nifty_fut():
        # Fetching spot and future data using Fyers quotes  NSE:NIFTY24OCTFUT
        data = {
            "symbols":  'NSE:NIFTY24OCTFUT'
        }
        response = fyers.quotes(data=data)
        
    # Nifty Spot data
        future_data = response  # Nifty Future data

        # print(future_data)
    
        return  future_data

    spot_data = fetch_nifty_spot()
    future_data = fetch_nifty_fut()

    spot_data_live= spot_data['d'][0]['v']['lp']
    future_data_live = future_data['d'][0]['v']['lp']
    # fatch percentage change
    def fetch_nifty_depth():
        # Fetching spot and future data using Fyers quotes  NSE:NIFTY24OCTFUT
        data = {
        "symbol":'NSE:NIFTY50-INDEX',
        "ohlcv_flag":"1"
        }

        response = fyers.depth(data=data)
        spot_data_depth = response  # Nifty Spot data
        # future_data = response['d'][1]  # Nifty Future data
        # print(spot_data)
        # print(spot_data_depth)
    
        return spot_data_depth # future_data
    spot_data_Dept=fetch_nifty_depth()

    # print(spot_data_Dept)

    # fatch percentage change
    def fetch_nifty_fut_depth():
        # Fetching spot and future data using Fyers quotes  NSE:NIFTY24OCTFUT
        data = {
        "symbol":'NSE:NIFTY24OCTFUT',
        "ohlcv_flag":"1"
        }

        response = fyers.depth(data=data)
        fut_data_depth = response  # Nifty Spot data
        # future_data = response['d'][1]  # Nifty Future data
        # print(spot_data)
        # print(spot_data_depth)
    
        return fut_data_depth # future_data
    fut_data_Dept=fetch_nifty_fut_depth()


    symbol = 'NSE:NIFTY50-INDEX'
    # fetch atm premiums
    def fetch_atm_premiums():
        quote = fyers.quotes({"symbols": symbol})
        print("Underlying Quote Response:", quote)  # Debugging
        if quote['s'] == 'ok':
            data = quote['d'][0]
            ltp = data['v']['lp']  # Last traded price
            
            # Assuming strike prices are rounded to the nearest 50
            strike_price = round(ltp / 50) * 50
            print(strike_price)
            
            # Correct format: DDMMMYYYY (e.g., 29AUG2024)
            expiry = '24O24'  # Use the correct expiry date "symbol":"NSE:NIFTY2451622000PE"
            call_symbol = f'NSE:NIFTY{expiry}{strike_price}CE'
            put_symbol = f'NSE:NIFTY{expiry}{strike_price}PE'

            # Fetch Call Premium
            call_quote = fyers.quotes({"symbols": call_symbol})
            print("Call Option Quote Response:", call_quote)  # Debugging
            call_premium = None
            if call_quote['s'] == 'ok' and 'd' in call_quote and 'lp' in call_quote['d'][0]['v']:
                call_premium = call_quote['d'][0]['v']['lp']
            else:
                print(f"Failed to fetch call premium for symbol {call_symbol}. Error: {call_quote}")

            # Fetch Put Premium
            put_quote = fyers.quotes({"symbols": put_symbol})
            print("Put Option Quote Response:", put_quote)  # Debugging
            put_premium = None
            if put_quote['s'] == 'ok' and 'd' in put_quote and 'lp' in put_quote['d'][0]['v']:
                put_premium = put_quote['d'][0]['v']['lp']
            else:
                print(f"Failed to fetch put premium for symbol {put_symbol}. Error: {put_quote}")

            return strike_price, call_premium, put_premium
        else:
            print("Failed to get underlying quote")
            return None, None, None

    strike_price, call_premium, put_premium = fetch_atm_premiums()

    spread = spot_data_live - future_data_live

    sheet.range('B5').color = (196,196,177)

    sheet['B3'].value = spot_data_live
    sheet['B4'].value = future_data_live
    sheet['B5'].value = spread
    sheet['C3'].value = spot_data['d'][0]['v']['ch']
    sheet['C4'].value = future_data['d'][0]['v']['ch']
    sheet['D3'].value = spot_data_Dept['d']['NSE:NIFTY50-INDEX']['chp']
    sheet['D4'].value = fut_data_Dept['d']['NSE:NIFTY24OCTFUT']['chp']
    sheet['E3'].value = f'Above' if spot_data['d'][0]['v']['lp'] > spot_data['d'][0]['v']['prev_close_price'] else 'Below'
    sheet['E4'].value = f'Above' if future_data['d'][0]['v']['lp'] > future_data['d'][0]['v']['prev_close_price'] else 'Below'
    if sheet['E3'].value == 'Above':
        sheet.range('B3:E3').color = (30,179,0)
    elif sheet['E3'].value == 'Below':
        sheet.range('B3:E3').color = (255,0,0)
    if sheet['E4'].value == 'Above':
        sheet.range('B4:E4').color = (30,179,0)
    elif sheet['E4'].value == 'Below':
        sheet.range('B4:E4').color = (255,0,0)
    # option premium
    sheet['G3'].value = strike_price
    sheet['H3'].value = call_premium
    sheet['I3'].value = put_premium
    sheet['J3'].value = call_premium + put_premium

    def fetchOHLC(ticker,interval,duration):
        """extracts historical data and outputs in the form of dataframe"""
        instrument = ticker
        data = {"symbol":instrument,"resolution":interval,"date_format":"1","range_from":dt.date.today()-dt.timedelta(duration),"range_to":dt.date.today(),"cont_flag":"1"}
        sdata=fyers.history(data)
        # print(sdata)
        sdata=pd.DataFrame(sdata['candles'])
        sdata.columns=['date','open','high','low','close','volume']
        sdata['date']=pd.to_datetime(sdata['date'], unit='s')
        # sdata.date=(sdata.date.dt.tz_localize('UTC').dt.tz_convert('Asia/Kolkata'))
        sdata['date'] = sdata['date'].dt.tz_localize(None)
        sdata=sdata.set_index('date')
        return sdata


    exchange='NSE'
    sec_type='INDEX'
    symbol='NIFTY50'
    ticker=f"{exchange}:{symbol}-{sec_type}"
    # print(ticker)

    Histrocal_data_spot=fetchOHLC(ticker,'D',366)
    # print(Histrocal_data_spot)

    start_row = 6

    # Loop through each EMA period and calculate the EMA, then write to Excel
    for i, ind_name in enumerate([10, 21, 63, 150, 200]):


        ema = ta.ema(Histrocal_data_spot['close'], ind_name)
        
        
        # Calculate the row number for each EMA, incrementing with each iteration
        row = start_row + i
        spread = Histrocal_data_spot['close'].iloc[-1]-ema
        perc_chang= (spread/Histrocal_data_spot['close'].iloc[-1])*100
        


        # Write the EMA name (e.g., 'EMA_10') in column A
        sheet.range(f'A{row}').value = f"EMA_{ind_name}"
        
        # Write the corresponding EMA value in column B
        sheet.range(f'B{row}').value = ema.iloc[-1]
        sheet.range(f'C{row}').value = spread.iloc[-1]
        sheet.range(f'D{row}').value = perc_chang.iloc[-1]
        if spread.iloc[-1]>0 :
            sheet.range(f'E{row}').value =f'Above EMA {ind_name}'
            sheet.range(f'B{row}:E{row}').color = (30,179,0)
        else:
            sheet.range(f'E{row}').value = f'Below EMA {ind_name}'
            sheet.range(f'B{row}:E{row}').color = (255,0,0)

    # fetch AVWAP
    def fetch_fut_OHLC_AVWAP(ticker,interval,duration):
        """extracts historical data and outputs in the form of dataframe"""
        instrument = ticker
        data = {"symbol":instrument,"resolution":interval,"date_format":"1","range_from":dt.date.today()-dt.timedelta(duration),"range_to":dt.date.today(),"cont_flag":"1"}
        sdata=fyers.history(data)
        # print(sdata)
        sdata=pd.DataFrame(sdata['candles'])
        sdata.columns=['date','open','high','low','close','volume']
        sdata['date']=pd.to_datetime(sdata['date'], unit='s')
        # sdata.date=(sdata.date.dt.tz_localize('UTC').dt.tz_convert('Asia/Kolkata'))
        sdata['date'] = sdata['date'].dt.tz_localize(None)
        sdata=sdata.set_index('date')
        return sdata


    # download data for AVWAP
    ticker= 'NSE:NIFTY24OCTFUT'
    data_fut_avwap=fetch_fut_OHLC_AVWAP(ticker,'D',188)
    # print(data_fut_avwap)

    def calculate_avwap(data_fut_avwap, anchor_index):

        # Calculate typical price
        data_fut_avwap['Typical_Price'] = (data_fut_avwap['high'] + data_fut_avwap['low'] + data_fut_avwap['close']) / 3

        # Multiply typical price by volume
        data_fut_avwap['Price_Volume'] = data_fut_avwap['Typical_Price'] * data_fut_avwap['volume']
        
        # Create slices of data starting from the anchor point
        # df_anchor = data_fut_avwap.iloc[anchor_index:].copy()

        # Cumulative sum of price * volume and volume
        data_fut_avwap['Cumulative_Price_Volume'] = data_fut_avwap['Price_Volume'].cumsum()
        data_fut_avwap['Cumulative_Volume'] = data_fut_avwap['volume'].cumsum()

        # Calculate AVWAP
        data_fut_avwap['AVWAP'] = data_fut_avwap['Cumulative_Price_Volume'] / data_fut_avwap['Cumulative_Volume']
        
        return data_fut_avwap['AVWAP']


    avwap = calculate_avwap(data_fut_avwap, anchor_index=0)
    print(avwap)
    sheet['B11'].value = avwap.iloc[-1]
    sheet['B12'].value = fut_data_Dept['d']['NSE:NIFTY24OCTFUT']['atp']

    Histrocal_data_spot['RSI'] = ta.rsi(Histrocal_data_spot['close'], length=14)
    # print(Histrocal_data_spot)
    sheet['B13'].value = Histrocal_data_spot['RSI'].iloc[-1]



    if Histrocal_data_spot['RSI'].iloc[-1]>70:
            sheet['E13'].value =   'Overbought'
            # sheet['B13'].color = (30,179,0)
            # sheet['E13'].color = (30,179,0)
    elif Histrocal_data_spot['RSI'].iloc[-1]<30:
            sheet['E13'].value = 'OverSold'
            sheet['B13'].color = (255,0,0)
            sheet['E13'].color = (255,0,0)
    elif Histrocal_data_spot['RSI'].iloc[-1]<=70 and Histrocal_data_spot['RSI'].iloc[-1]>=30:
            sheet['E13'].value = 'Neutral'
            sheet['B13'].color = (143,133,0)
            sheet['E13'].color = (143,133,0)

    Histrocal_data_spot['ATR'] = ta.atr(Histrocal_data_spot['high'], Histrocal_data_spot['low'], Histrocal_data_spot['close'], length=14)
    sheet['B14'].value = Histrocal_data_spot['ATR'].iloc[-1]
    sheet.range('B14').color = (196,196,177)

        # 52-week, weekly, and monthly highs/lows
    ticker_symbol = '^NSEI'

    # Fetch the stock data
    ticker = yf.Ticker(ticker_symbol)

    # Get the stock information (including 52-week high/low)
    stock_info = ticker.info

    # Extract the 52-week high and low from the info dictionary
    fifty_two_week_high = stock_info.get('fiftyTwoWeekHigh')
    fifty_two_week_low = stock_info.get('fiftyTwoWeekLow')

    # if sheet['E3'].value == 'Above':
    #     sheet.range('B3:E3').color = (30,179,0)
    # elif sheet['E3'].value == 'Below':
    #     sheet.range('B3:E3').color = (255,255,0)
    Histrocal_data_spot['Weekly_High'] = Histrocal_data_spot['high'].rolling(window=7).max()
    Histrocal_data_spot['Weekly_Low'] = Histrocal_data_spot['low'].rolling(window=7).min()
    Histrocal_data_spot['Monthly_High'] = Histrocal_data_spot['high'].rolling(window=30).max()
    Histrocal_data_spot['Monthly_Low'] = Histrocal_data_spot['low'].rolling(window=30).min()
    # print(Histrocal_data_spot)
    sheet['B15'].value = Histrocal_data_spot['Weekly_High'].iloc[-1]
    sheet['C15'].value = Histrocal_data_spot['Weekly_Low'].iloc[-1]
    sheet['B16'].value = Histrocal_data_spot['Monthly_High'].iloc[-1]
    sheet['C16'].value = Histrocal_data_spot['Monthly_Low'].iloc[-1]
    sheet['B17'].value = fifty_two_week_high
    sheet.range('B17').color = (196,196,177)
    sheet.range('C17').color = (196,196,177)
    sheet['C17'].value = fifty_two_week_low


# while True:
    
#     time.sleep(20)

# print(data1)

while True:
    process_and_export_to_excel()
    time.sleep(2)