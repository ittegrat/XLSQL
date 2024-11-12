# XLSQL
The XLSQL add-in lets you:
  - use the SQL language to query and filter data in Excel ranges
  - read, write, and update SQLite databases from Excel

## Worksheet Functions
 - **SQL.OpenConnection**: open a connection to a SQLite database.
 - **SQL.CloseConnection**: close a SQLite database connection.
 - **SQL.ListConnections**: list open connections whose name matches a regex pattern.

 - **SQL.Query**: run a query and returns a result set.
 - **SQL.Execute**: execute a non-query statement and returns the number of rows affected by it.

 - **SQL.QueryRange**: run a query on multiple ranges and return a result set.

 - **SQL.CreateTable**: create a virtual table backed by a range in 'temp' schema.
 - **SQL.FreezeTable**: freeze a virtual table set up using 'CreateTable'.
 - **SQL.UnfreezeTable**: unfreeze a virtual table set up using 'CreateTable'.
 - **SQL.RefreshTable**: refresh a frozen virtual table set up using 'CreateTable'.
 - **SQL.ListTables**: list virtual tables whose name matches a regex pattern.

 - **SQL.CreateQuery**: create a reusable, possibly parameterized, query.
 - **SQL.DeleteQuery**: delete a reusable query.
 - **SQL.ListQueries**: list reusable queries whose name or sql text matches the corresponding regex pattern.


## Macro Functions
 - **XLSQL.CreateDatabase**: create a SQLite database file and open a connection to it.
 - **XLSQL.ShowTab**: show the XLSQL ribbon tab.
 - **XLSQL.HideTab**: hide the XLSQL ribbon tab.

## Ribbon Functions
 - **XLSQL.OpenFileDB**: open a connection to an existent SQLite database.
 - **XLSQL.NewMemoryDB**: create a new memory database and opens a connection to it.
 - **XLSQL.NewFileDB**: create a new file database and opens a connection to it.
 - **XLSQL.CloseDB**: close one or more SQLite database connections.

 - **XLSQL.NewRangeTable**: create a virtual table backed by a range in 'temp' schema.
 - **XLSQL.FreezeRangeTable**: freeze the values in a 'Range' table.
 - **XLSQL.UnfreezeRangeTable**: unfreeze the values in a 'Range' table.
 - **XLSQL.RefreshRangeTable**: update the values in a frozen 'Range' table.
