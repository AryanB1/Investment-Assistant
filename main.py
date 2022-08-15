# Libraries
import yfinance as yf
import pyodbc as db

# Check For drivers
required_drivers = [x for x in db.drivers() if 'ACCESS' in x.upper()]
print(f'MS-Access Drivers : {required_drivers}')
# Get Stock Tickers
try:
    # Connection to Database
    connection_msg = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\abhar\OneDrive\Documents\pythondb.accdb;'
    connection = db.connect(connection_msg)
    print("Connected To Database")
    # Object to begin SQL Injections
    cursor = connection.cursor()

    # Get Ticker Symbols from SQL database
    rows = cursor.execute("select Field1 from Table1").fetchall()
    # List to store stock tickers
    stocks = []
    # Iterated Value in While Loop
    x = 0
    # Adds every value in the database to tickers list
    while True:
        try:
            stocks.append(rows[x][0])
            x += 1
        except IndexError:
            break
    # Iterated Value in for loop
    y = 0
    for stock in stocks:
        y += 1
        ticker_yahoo = yf.Ticker(stock)
        data = ticker_yahoo.history()
        last_quote = data['Close'].iloc[-1]
        cursor.execute("UPDATE Table1 SET Field2=? WHERE ID=? ", (last_quote, str(y)))
        connection.commit()
    print('Data Inserted')
except db.Error as error:
    print("Error in Connection", error)
