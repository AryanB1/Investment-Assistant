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
- yfinance library in python [docs](https://pypi.org/project/yfinance/)
- pyodbc [docs](https://pypi.org/project/pyodbc/)
## Disclaimer
This code is not professional advice, and I am not a professional in the field of finance. Please consult with a financial advisor before making any decisions. This software provides no warranty to the user. By using this software you are entirely liable for any damages that it causes. 

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
