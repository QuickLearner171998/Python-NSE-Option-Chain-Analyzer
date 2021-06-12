# Python NSE-Option-Chain-Analyzer

For doing technical analysis for option traders, the Option Chain is the most important tool for deciding entry and exit
strategies. The National Stock Exchange (NSE) has a website which displays the option chain for traders in near
real-time. This program retrieves this data from the NSE site and then generates useful analysis of the Option Pairing for
the specified Index or Stock. It also continuously refreshes the Option data.


## Disclaimer:

> #### This software is an unofficial software and is not affiliated with / endorsed or approved by the National Stock Exchange (NSE) or Mr. Sameer Dharaskar in any way
>
>#### This is purely an enthusiast program intended for educational purposes only and is not financial advice
>
>#### By downloading this software you acknowledge that you are using this at your own risk and that I am is not responsible for any damages that may occur to you due to the usage or installation of this program
>
>#### All NSE name/symbols are owned by the National Stock Exchange

## Installation:

Any one of Method1 or Method2 can be used to run the program.
1. Method 1
    - Directly run `NSE_Option_Chain_Analyzer.exe`



2. Method 2
    - Install missing modules using `pip install -r requirements.txt`
    - `python NSE_Option_Chain_Analyzer.py`

## Usage:

1. Set Index Mode or Stock Mode

2. Select your Index or Stock

3. Select the Expiry Dates

4. Enter the strike prices and the respective Buy/Sell prices

5. Set the interval you want the program to refresh (Optional : Defaults to 1 minute)

6. Click Start

## Note:

- If there is an error in fetching dates on login screen then try refreshing

- If there is an error in fetching dates on main screen then try stopping and again starting from option menu

- If you face any issue or have a suggestion then feel free to open an issue.

- It is recommended to enable logging and then send the `NSE-OCA.log` file or the console output for reporting issues

- In case of network or connection errors the program doesn't crash and will keep retrying until manually stopped

- If a ZeroDivisionError occurs or some data doesn't exist the value of the variable will be defaulted to 0

- Set `load_nse_icon` option to `False` in the configuration file to prevent downloading the NSE icon in the `.py`
  version to speed up loading time

- The executable file is tested on Windows 8 and Windows 10.

## Dependencies:

- [pyinstaller](https://pypi.org/project/pyinstaller/) is used for compiling the program to a .exe file

- [beautifulsoup4](https://pypi.org/project/beautifulsoup4/) is used for scraping the list of stocks and indices

- [pandas](https://pypi.org/project/pandas/) is used for storing and manipulating the data

- [requests](https://pypi.org/project/requests/) is used for accessing and retrieving data from the NSE website

- [stream-to-logger](https://pypi.org/project/streamtologger/) is used for debug logging

- [tksheet](https://pypi.org/project/tksheet/) is used for the table containing the data

- [win10toast](https://pypi.org/project/win10toast/) is used for Windows 10 Toast notifications

## Features:

- The program continuously retrieves and refreshes the option chain giving near real-time analysis to the traders

- New data rows are added only if the NSE server updates its time or data (To prevent displaying duplicate data)

- Supported Indices and
  Stocks: https://www.nseindia.com/products-services/equity-derivatives-list-underlyings-information

- Supports multiple instances with different indices/stocks and/or strike prices selected

- Red and Green colour indication for data based on trends

- Stop and Start manually

- You can select all table data using Ctrl+A or select individual cells, rows and columns

- Then you can copy it using Ctrl+C or right click menu

- You can then paste it in any spreadsheet application (Tested with MS Excel and Google Sheets)

- Export table data to `.csv` file

- Real time exporting data rows to `.csv` file

- Dumping entire Option Chain data to a `.csv` file

- Auto stop the program at 3:30pm when the market closes

- Alert if the last time the data from the server was updated is 5 minutes or more

- Auto Checking for updates

- Debug Logging

- Saves certain settings in a configuration file for subsequent runs. Saved Settings:
    * Load App Icon
    * Index/Stock Mode
    * Selected Index
    * Selected Stock
    * Refresh Interval
    * Live Export
    * Notifications
    * Dump entire Option Chain
    * Auto stop at 3:30pm
    * Warn Late Server Updates
    * Auto Check for Updates
    * Debug Logging

- Keyboard shortcuts for all options
