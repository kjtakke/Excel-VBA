Dim conn As Object, rst As Object

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

' OPEN CONNECTION
conn.Open "DRIVER=SQLite3 ODBC Driver;Database=C:\Path\To\SQLite\Database.db;"

strSQL = "SELECT Vessels.vsl_name, Vessels.dwt FROM Vessels GROUP BY Vessels.vsl_name, Vessels.dwt ORDER BY Vessels.vsl_name ;"

' OPEN RECORDSET
rst.Open strSQL, conn, 1, 1

' OUTPUT TO WORKSHEET
Worksheets("results").Range("A1").CopyFromRecordset rst
rst.Close

' FREE RESOURCES
Set rst = Nothing: Set conn = Nothing
