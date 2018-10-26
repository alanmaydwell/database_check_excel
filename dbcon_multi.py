#!/usr/bin/env python
"""
Create and manage database connections, run SQL queries and extract results.
Can use either cx_Oracle or pyodbc to make the connection

"""
from __future__ import print_function
import time

# Handling for either/or import situaiton as we
# don't necessarily need both pyodbc and cx_Oracle
# Also exception names we later want to handle depend on success
# of these imports
FAILED_IMPORTS = []
DB_EXCEPTIONS = []

try:
    import pyodbc
except ImportError as err:
    FAILED_IMPORTS.append("pyodbc")
else:
    DB_EXCEPTIONS.append(pyodbc.Error)

try:
    import cx_Oracle
    from cx_Oracle import DatabaseError
except ImportError as err:
    FAILED_IMPORTS.append("cx_Oracle")
else:
    DB_EXCEPTIONS.append(DatabaseError)

if FAILED_IMPORTS == ["pyodbc", "cx_Oracle"]:
    print("Critical Failure. Failed to import db module.")


class DbCon(object):
    def __init__(self, username, password, database,
                 odbc_driver="", do_nothing=False):
        """
        Create database connection using either pyodbc or cx_Oracle
        (depending on odbc_driver param).
        Has method for exectuing SQL query.
        Result of query stored in self.results
        Errors held in self.errors
        Args:
            username (str) - username
            password (str) - password
            database (str) - database name
            odbc_driver (str) - Optional odbc driver name
                                e.g. "Oracle in Oraclient11g_home"
                                or "Oracle in Instantclient11_1"
                                If None/empty cx_Oracle connection will be
                                used instead of ODBC.
            do_nothing (bool) - don't automatically make connection if true
        """
        #Connection type:
        if odbc_driver:
            self.db_module = pyodbc
        else:
            self.db_module = cx_Oracle

        #Holds error messages
        self.errors = []
        #Holds results
        self.results = []
        #Holds column headings
        self.headings = []
        #Execution time for query as date/time string (updated by self.runsql()
        self.execution_time = ""
        self.database = database

        # Construct connection string
        # ODBC connection string (depends on tnsnames.ora)
        if odbc_driver:
            self.constring = "Driver={%s};Dbq=%s;Uid=%s;Pwd=%s" % (odbc_driver, database, username, password)
            ##print(">>>", odbc_driver)
        #cx_Oracle connection string
        else:
            # Direct connection - expects !<database name>,sid
            # e.g. "!dbabc,hub"
            if database.startswith("!") and "," in database:
                parts = database.split(",")
                host = parts[0][1:].strip()
                sid = parts[1].strip()

                # Construct direct connection string
                self.constring = username + "/" + password + "@" + host + ":1521/" + sid
            # tns type connection string
            # note don't really need these for simple connections like this as
            # could use -  self.cnxn = cx_Oracle.connect(username, password, database)
            else:
                self.constring = username + "/" + password + "@" + database

        #Sometimes we might not want to automatically open the connection
        if not do_nothing:
            self.open()

    def open(self):
        """Open and test database connection"""
        #Try to make database connection using connection string
        try:
            self.cnxn = self.db_module.connect(self.constring)
        # DatabaseError - cx_Oracle, pyodbc.Error - pyodbc
        except (DatabaseError, pyodbc.Error) as err:
            self.cnxn = None
            self.errors.append(str(err))

    def close(self):
        """If connection exists, close it"""
        if self.cnxn:
            self.cnxn.close()

    def runsql(self, sql, params=()):
        """Execute SQL using current connection, retrieve results and 
        store in self.results but only if there's a current connection
        Args:
            sql (str) - sql to be executed
            params - optional container of sql substitution parameters
        """
        self.execution_time = time.strftime("%d-%b-%Y %H:%M:%S")
        #Don't run if there's no connection
        if not self.cnxn:
            self.errors.append("Can't execute SQL because no connection.")
        else:
            self.results, self.headings, self.errors = self.execute(sql, params)

    def execute(self, sql, params=()):
        """Execute sql using current connection and retrieve results
        Args:
            sql - sql to execute
            params - optional subsitution parameters if format valid for cx_Oracle
        Returns:
            SQL query result (list of tuples)
            Column headings (list of strings)
            Error messages (list of strings)
        """
        #Create cursor and execute SQL;
        local_errors = []
        headings = []
        rows = []
        cursor = self.cnxn.cursor()
        try:
            if not params:
                cursor.execute(sql)
            else:
                cursor.execute(sql, params)
        except Exception as err:
            local_errors.append("Error on execution:" + str(err))
        else:
            #SQL exececution successful - retrieve results
            try:
                rows = cursor.fetchall()
            except Exception as err:
                local_errors.append("Error on fetching results:" + str(err))
            else:
                #Also capture column headings
                headings = [d[0] for d in cursor.description]
        return rows, headings, local_errors

    def db_info(self):
        """Get some info from v$database
        Return details if found
        """
        if self.cnxn:
            db_info_sql = "select name, db_unique_name, dbid, created from v$database"
            rows, headings, errors = self.execute(db_info_sql)
            if not errors:
                result = {h:str(r) for h, r in zip(headings, rows[0])}
            else:
                result = errors
        else:
            result = "Can't retrieve details as no connection"
        return result


# Example/test connection
if __name__ == "__main__":
    database = "!xxxxxxxxxxxxxxxxxxxx.eu-west-2.rds.amazonaws.com,hub"
    username = raw_input("Username:")
    password = raw_input("Password (will show!):")

    # Query with no paramters
    test_con = DbCon(username, password, database)
    test_con.db_info()
    print(test_con.results)
    test_con.runsql("select sysdate from dual")
    test_con.close()
    print("Results:", test_con.results)
    print("Errors:", test_con.errors)
    print("finished")

    # Query with parameter
    test_con = DbCon(username, password, database)
    test_con.db_info()
    print(test_con.results)
    test_con.runsql("select * from hub_processes where process_code = :pc", ["BEW01"])
    test_con.close()
    print("Results:", test_con.results)
    print("Errors:", test_con.errors)
    print("finished")
