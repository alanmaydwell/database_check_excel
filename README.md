# Database Checker
- Runs multiple SQL queries using connection and query details extracted from a specially formatted spreadsheet.

- Writes the results back to a copy of the spreadsheet with auto-generated filename (including time/date) in Results subfolder below the parent folder.

- Can have simple conditional highlighting of the results.


Requires

- Python 2.7
- openpyxl
- pyodbc or cx_Oracle

## Files

### database_check_excel.py
- Execute this to run the checks
- Spreadsheet(s) to be included can be specified in the __main__ block near the end of the script or passed as command-line arguments. Defaults to queries.xlsx.
- If using odbc connection, the script will auto-choose the last driver returned by pyodbc.drivers(). Alternatively the desired choice can be hard-coded near the end of the script.

### queries.xlsx
- Example spreadsheet. 
- Database details in this file are not real.

### dbcon_multi.py
Contains class used to make the database connections and run the queries.