# Libraries
import yfinance as yf
import pyodbc as db
import matplotlib.pyplot as plt
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
    rows = cursor.execute("select Ticker from Table1").fetchall()
    # List to store stock tickers
    stocks = []
    stock_vals = []
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
        stock_vals.append(last_quote)
        cursor.execute("UPDATE Table1 SET currentPrice=? WHERE ID=? ", (round(last_quote, 2), str(y)))
        connection.commit()
    date = cursor.execute("select whenBought from Table1").fetchall()
    # Graph Stocks
    graph_data = yf.download(stocks,'2016-01-01')
    graph_data['Adj Close'].plot()
    plt.show()
    # List to store stock tickers
    dates = []
    # Iterated Value in While Loop
    x = 0
    # Adds every value in the database to tickers list
    while True:
        try:
            dates.append(date[x][0])
            x += 1
        except IndexError:
            break
    y = 0
    for date in dates:
        try:
            name = stocks[y]
            stonk = yf.download(name, date)
            stonk = float(stonk["Close"][0])
            change = ((float(stock_vals[y]) - stonk)/(abs(stonk)))*100
            y += 1
            cursor.execute("UPDATE Table1 SET percentChange=? WHERE ID=? ", (round(change, 2), str(y)))
            connection.commit()
        except TypeError:
            break
    print('Data Inserted')
except db.Error as error:
    print("Error in Connection", error)
