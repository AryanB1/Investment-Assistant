# Yfinance-to-MS-database
## How to setup
- First, find the following code snippit on line eleven of the main.py file 
```
connection_msg = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\abhar\OneDrive\Documents\pythondb.accdb;'
```
- The ```C:\Users\abhar\OneDrive\Documents\pythondb.accdb``` substring must be changed to the file path of your database
- Furthermore, the fields and tables must be edited to match the name of your fields and tables. For simplicity, I left the initial values of these fields as the default values assigned by Microsoft
- Finally, the Code on line 26 is preset to return the latest price of the stock, visit the yfinance documentation linked below to change this 
## Requirements
- yfinance library in python docs
- 
