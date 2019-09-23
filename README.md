# database_check_excel.py

## Requires

- Python 3 or Python 2 (Python 3 recommended)
- Oracle Client
- openpyxl Python module
- cx_Oracle Python module and/or pyodbc Python module (as minimum need one of these two)
- In addition, if using pyodbc then an odbc driver is needed (e.g. Oracle odbc driver, which is optional add-on to Oracle client. Note 32-bit driver needed for 32-bit Python, 64-bit driver for 64-bit Python.)

A spreadsheet application is also needed to setup data and view the results. Excel or LibreOffice can be used for this purpose. 

## Purpose
    
- Reads and executes multiple SQL queries from a specially formatted Excel spreadsheet.
Writes the results of the queries back to a new copy of the spreadsheet with time/date in filename. These are saved in **results** sub-folder, which will be created if it does not exist.

- The queries can be against different databases - the database details are also stored within the same spreadsheet.

- The results of a query can either be written to a single spreadsheet cell or to multiple cells in a designated position in on a particular spreadsheet tab.

- It is also possible to specify conditional highlighting to a single-cell result or to multiple-cell results. This highlighting uses Python expressions stored within the spreadsheet and does not rely on any Excel conditional highlighting functionality (although this could also potentially be applied).

- The database queries are run in parallel, with main aim of preventing connection/timeout problems with one database delaying execution of queries against a different database.

- Multiple spreadsheets can be processed in a single run.

- If connection errors are encountered the script should still continue and related error messages will be written to the results spreadsheet.

- It is optionally possible to have the results written back to the original spreadsheet too.

## How to run

> **Aside concerning means used to connect to the databases**.        
> This script can use either pyodbc or cx_Oracle to make database connections but there are quirks to either approach.     
> An earlier version of this script relied on only odbc to make the database connection and used database names corresponding with the tnsnames.ora file.        
> Direct connections are now also covered but the database name needs to be prefixed with an exclamation mark to indicate a direct connection is being used, as described in Column Details section below.      
> **Note**
> - If using an ODBC driver this needs to be specified in the `if __name__ == "__main__"` block at the end of database_check_excel.py. There are some commented out examples there already (including simple auto-detection).        
> - cx_Oracle seems only to work with direct connections, so need to specify database names in "!" format. 
> - Conversely Oracle ODBC drivers only seem to work with tnsnames.ora connections. Although Microsoft ODBC drivers may work with either direct or tnsnames connections.
> - Nevertheless, this script is setup on the basis that cx_Oracle is for direct connections and ODBC for tnsnames.ora connections.


### (1) Setup standard spreadsheet  
- Create standard spreadsheet with required details, or or use an existing one. There are several examples in this repo.    
- See **Spreadsheet Setup** section below for details about how to setup the spreadsheet.
- Place the spreadsheet(s) in the same folder as database_check_excel.py


### (2) Run database_check_excel.py 
Run database_check_excel.py in whatever way is possible depending on Python setup. Can be run from command line or via file manager/gui or from within a Python IDE (although this may break password masking)

The spreadsheets to be included in a run can be specified in two ways:

(1) In filenames list in the `if __name__ == "__main__":` block near the end of the script, e.g.

```python
if __name__ == "__main__":

    filenames = ['hub_user_admin_state.xlsx', 'hub_check(dev_test).xlsx']
```    
 
(2) As command-line arguments, e.g. `python database_check_excel.py hub_user_admin_state.xlsx hub_check(dev_test).xlsx`    
When filenames are supplied at command-line, the filenames specified within the script are ignored.    


## The Results
- When finished, a new spreadsheet should have been created in the **results** folder below the folder the script was run from.
- If database connection errors are encounterd the script should still complete and related error messages recorded in the generated spreadsheet.
- The results spreadsheets have a yellow background applied to their topmost rows. This is to help distinguish them from the original spreadsheets.
- To aid navigation the tab names in column A of the summary tab are hyperlinks to their respective tabs. Cell B2 in the non-Summary tabs contains a link back to the summary.


## Spreadsheet Setup
Details of database connections and SQL queries are defined in one or more **"Database Query"** tabs. Note these are not automatically included in a run but have to be listed in the **Run Tab**. Results are recorded in a single cell in the **"Database Query"** tab and optionally can be tabulated to a specified region in a **"Results"** tab.    
   
There are some example spreadsheets within this repo.

### Summary Tab
- Holds summary of results at end of run.
- Completed during run. Tab names in coloumn A are hyperlinks to corresponding tabs.
- Not meant to be updated manually. Doesn't need to be in the parent spreadsheet. Is created afresh when script run.

### Run Tab
- Column B on the Run Tab is used to specify the tab names of the tabs to be included when the spreadsheet is processed. Values are read from rows 5 to 20, so a currently a limit of 15 Query Tabs per spreadsheet.
- Cell D5 is used to control whether the original spreadsheet will be updated by the run. Set anything starting "Y" or "y" in D5 for this to happen.

### Database Query Tabs
> Note the script identifies each standard column by particular headings ("Username", "Password", "Database" etc) in row 6 and examines columns A to T     
> Accordingly:    
> (a) Do not rename these headings unless you make corresponding changes within the Python script itself.    
> (b) It's fine to change the order of the columns as long as the expected headings are in row 6 between columns A and T.

#### Range Used During Run
- The rows processed within a particular query tab are defined by the **start row** and **end row** values stored respectively in cells C3 and C4.     
- C4 value cannot be lower than C3 - no provision for reverse order!     
- The rows included within a run can be further restricted through the use of the Skip column (see below).

#### Column Details
The table below describes how each column on a Query Tab is used.

| Column Name | Mandatory? | Description |
| --- | --- | --- |
| Skip | No | To skip a row put anything starting "Y" or "y"|
| Username | Yes | Database username |
| Password | No | Database Password. If a blank value is encountered, the script will request a single password. Becomes mandatory if you want to use more than one pasword. |
|Database| Yes | (a) If using pyodbc and tnsnames.ora, needs just the database name. (b) If using cx_Oracle and direct connection, "!", then database name, then comma, then sid (e.g. !rds.hub.aws.tst.legalservices.gov.uk,hub). "!" used to denote direct connection, port 1521 used. |
| Heading | No | Heading text for the query. Can be used just to aid identification. Automatically copied to any related results tab. |
| SQL | Yes | SQL query to run. |
| Results Tab | No | Optional name of results tab. When set, results will be tabulated in specified tab. When used, also need to set Results Column letter. |
| Results Column | No | Lefthand column letter to write results to on results tab. Required if Results Tab set |
| Results Row | No | Topmost row number to write results to on results tab. Only relevant if Results Tab specified. Defaults to row 6 (matches SpreadsheetRun.heading_row value).
| Results Condition | No | Optional row-based condition applied to results in results tab. Condition is a Python expression. Variable c represents column number (starting with 1), x represents the cell value. This is a negative condition - "bad" highlight when true. |
| Local Condition | No | Optional condition applied to the "local result" (the whole query result written to the Result column). Condition is a Python expression. Variable x represents the result. Unlike Results Condition, this is a positive condition - "good" highlight when true. |
| Result | n/a | Script writes the results of the query to this cell (even when Results Tab specified). Background will be highlighted in accordance with any associated Local Condition. Multi row/column results are converted to comma-separated string. *Possibly a large volumn of data may break the Excel file.*|
| Date/Time | n/a | Script writes date/time here when recording results. | 

### Results Tabs
- If Results tabs were specified in any of the database query tabs then they can be added.
- This is not esssential as the script will create any results tabs it needs if they don't exist. However, this sometimes leads to corruption in the results spreadsheets, so it's best to create the results tabs in advance.
- Another advantage of creating results tabs in advance is it allows column widths to be set to suit the returned data.

### Colour Highlighting

> Seems markdown isn't good a colour handling.      
> To see examples of the actual colours look at column F in the Run tab of the example spreadsheets.

Colour highlighting is applied to the Result column in each Query tab and to rows in any Results tabs. The meaning of each colour is listed below.

#### Green
"Good" result. No problems with database connection or SQL query. Any Local Condition or Results Condition check also had "good" outcome.

#### Red
Fundamental fatal error such as failure to connect to the database or error on SQL execution.

#### Orange
Only occurs when Local Condition or Results Condition check appplied. Indicates "bad" result.

#### Purple
Only occurs when Local Condition or Results Condition check appplied. Indicates an exception was raised when trying to apply the condition. Likely error in the condition or incompatible data encountered.

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


