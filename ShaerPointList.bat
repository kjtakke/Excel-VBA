Option Explicit

Sub SQL_Editor()

SP_SQL.Show

End Sub

Sub Clr()
Worksheets("SharePoint_Data").Range("B2:ZZ99999").Value = ""
End Sub
Sub Svd()
Worksheets("Saved_Queries").Activate
End Sub
Sub Bk()
Worksheets("SharePoint_Data").Activate
End Sub
Sub Pk()
OpenSharePointItem.Show
End Sub



Sub SP_List()
Worksheets("Temp_Data").Range("B2:ZZ99999").Value = ""
Dim cnt As ADODB.Connection
Dim rst As ADODB.Recordset
Dim MySQL As String
Dim SITE As String
Dim LISTID As String
Dim SQL_SELECT As String
Dim SQL_FROM As String
Dim SQL_WHERE As String
Dim SQL_GROUPBY As String

On Error GoTo En:
SQL_SELECT = Worksheets("BaseData").Range("B3").Value
SQL_FROM = Worksheets("BaseData").Range("B4").Value
SQL_WHERE = Worksheets("BaseData").Range("B5").Value
SQL_GROUPBY = Worksheets("BaseData").Range("B6").Value

SITE = Worksheets("BaseData").Range("B1").Value
LISTID = Worksheets("BaseData").Range("B2").Value

Set cnt = New ADODB.Connection
Set rst = New ADODB.Recordset


'SQL Query
If SQL_WHERE = "" Then
    MySQL = "SELECT " & SQL_SELECT & " FROM " & SQL_FROM
Else
    MySQL = "SELECT " & SQL_SELECT & " FROM " & SQL_FROM & " WHERE " & SQL_WHERE
End If

If SQL_GROUPBY = "" Then
    MySQL = MySQL & ";"
Else
    MySQL = MySQL & " GROUP BY " & SQL_GROUPBY & ";"
End If

'MySQL = "SELECT [Request Military Assistance - New].Priority FROM [Request Military Assistance - New] GROUP BY [Request Military Assistance - New].Priority;"
'Debug.Print (MySQL)

'SharePoint Site and List
With cnt
    .ConnectionString = _
    "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes;DATABASE=" & SITE & ";LIST={" & LISTID & "};"
    .Open
End With

'Open session
rst.Open MySQL, cnt, adOpenForwardOnly, adLockReadOnly


    'List data location
    Dim arry As Variant
    If Not (rst.BOF And rst.EOF) Then
        Worksheets("Temp_Data").Range("B2").CopyFromRecordset rst
        arry = rst.GetRows
    End If
    
    
    
    'Send SharePoint data to an Array
    ReDim SP_List(0 To UBound(arry, 2), 0 To UBound(arry)) As Variant
    Dim Counter As Long, Internalcounter As Long
    'Unpivot Array
    For Counter = 0 To UBound(arry)
        For Internalcounter = 0 To UBound(arry, 2)
            SP_List(Internalcounter, Counter) = arry(Counter, Internalcounter)
        Next Internalcounter
    Next Counter



'End session
If CBool(rst.State And adStateOpen) = True Then rst.Close
Set rst = Nothing
If CBool(cnt.State And adStateOpen) = True Then cnt.Close
Set cnt = Nothing
En:
End Sub




Sub SP_List2()
Worksheets("SharePoint_Data").Range("B2:ZZ99999").Value = ""
Dim cnt As ADODB.Connection
Dim rst As ADODB.Recordset
Dim MySQL As String
Dim SITE As String
Dim LISTID As String
Dim SQL_SELECT As String
Dim SQL_FROM As String
Dim SQL_WHERE As String
Dim SQL_GROUPBY As String

On Error GoTo En:
SQL_SELECT = Worksheets("BaseData").Range("B3").Value
SQL_FROM = Worksheets("BaseData").Range("B4").Value
SQL_WHERE = Worksheets("BaseData").Range("B5").Value
SQL_GROUPBY = Worksheets("BaseData").Range("B6").Value

SITE = Worksheets("BaseData").Range("B1").Value
LISTID = Worksheets("BaseData").Range("B2").Value

Set cnt = New ADODB.Connection
Set rst = New ADODB.Recordset


'SQL Query
If SQL_WHERE = "" Then
    MySQL = "SELECT " & SQL_SELECT & " FROM " & SQL_FROM
Else
    MySQL = "SELECT " & SQL_SELECT & " FROM " & SQL_FROM & " WHERE " & SQL_WHERE
End If

If SQL_GROUPBY = "" Then
    MySQL = MySQL & ";"
Else
    MySQL = MySQL & " GROUP BY " & SQL_GROUPBY & ";"
End If

'MySQL = "SELECT [Request Military Assistance - New].Priority FROM [Request Military Assistance - New] GROUP BY [Request Military Assistance - New].Priority;"
'Debug.Print (MySQL)

'SharePoint Site and List
With cnt
    .ConnectionString = _
    "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes;DATABASE=" & SITE & ";LIST={" & LISTID & "};"
    .Open
End With

'Open session
rst.Open MySQL, cnt, adOpenForwardOnly, adLockReadOnly


    'List data location
    Dim arry As Variant
    If Not (rst.BOF And rst.EOF) Then
        Worksheets("SharePoint_Data").Range("B2").CopyFromRecordset rst
        arry = rst.GetRows
    End If
    
    
    
    'Send SharePoint data to an Array
    ReDim SP_List(0 To UBound(arry, 2), 0 To UBound(arry)) As Variant
    Dim Counter As Long, Internalcounter As Long
    'Unpivot Array
    For Counter = 0 To UBound(arry)
        For Internalcounter = 0 To UBound(arry, 2)
            SP_List(Internalcounter, Counter) = arry(Counter, Internalcounter)
        Next Internalcounter
    Next Counter



'End session
If CBool(rst.State And adStateOpen) = True Then rst.Close
Set rst = Nothing
If CBool(cnt.State And adStateOpen) = True Then cnt.Close
Set cnt = Nothing

With Worksheets("SharePoint_Data").Cells
    .RowHeight = 15
    .EntireColumn.AutoFit
End With
    Worksheets("SharePoint_Data").Rows("1:1").RowHeight = 43 = 43.5

En:
End Sub

