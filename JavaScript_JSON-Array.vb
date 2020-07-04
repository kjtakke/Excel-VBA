'INSTRUCTIONS AND EXAMPLE#################################################################################

'DATA
'    A     B     C
'1  Name   Qty  Age
'2  Kris   15   10
'3  Tim    30   15
'4  Jane   9    80

'CELL INPUT
'=JS_JSONcreator("Sheet1","A2","C2","myData","A1")

'OUTPUT
'const variableName = [
'       {
'             Name: 'Kris',
'             Qty: '15',
'             Age: '10',
'       },
'       {
'             Name: 'Tim',
'             Qty: '30',
'             Age: '15',
'       },
'       {
'             Name: 'Jane',
'             Qty: '9',
'             Age: '80',
'       },
'];




'SUBROUTINES AND FUNCTIONS##################################################################################

'MAIN SUBROUTINE
Sub WriteJS_JSON()
    'This Sub sends the JavaScript Array/JSON as a string to be written
    
    Call Writefile(JS_JSONcreator("Sheet1", "A3", "Z3", "variableName", "A2"), "My_JS_Array_JSON", ".js")
End Sub


'WRITE TO FILE
Sub Writefile(myTxt As String, fileName As String, fileExt As String)
    'This Sub write a file to your Desktop
    'myTxt              is the string of text to be written
    'fileName           is the name of the file to be written
    'fileExt            is the extent with "." of the file to be written
        
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("C:\Users\" & Environ("userName") & "\Desktop\" & fileName & fileExt, True)
    a.WriteLine myTxt
    a.Close
End Sub
    
        
'CREATE JAVASCRIPT ARRAY/JSON       
Function JS_JSONcreator(worksheet As String, topLeftCell As String, topRightCell As String, VariableName As String, headerTopLeftCell As String) As String
    'This Function creates a string of text in JavaScript Array (JSON) Format
    'worksheet          is the Excel worksheet that the data is in
    'topLeftCell        is the top left cell address less the sheet name below the table header
    'topRightCell       is the top right cell address less the sheet name below the table header
    'VariableName       is JavaScript Variable Name
    'headerTopLeftCell  is the far left header cell address less the sheet name
    
    Dim data, headings As Variant
    Dim ws As worksheet
    Dim headingNospace As String

    Set ws = Worksheets(worksheet)
    data = ws.Range(topLeftCell & ":" & topRightCell, ws.Range(topLeftCell & ":" & topRightCell).End(xlDown)).Value
    headings = ws.Range(headerTopLeftCell, ws.Range(headerTopLeftCell).End(xlToRight)).Value

    JS_JSONcreator = "const " & Replace(VariableName, " ", "_", , , vbTextCompare) & " = [" & vbNewLine
    For i = 1 To UBound(data)
        JS_JSONcreator = JS_JSONcreator & "       " & "{" & vbNewLine
        For j = 1 To DimentionCounter(headings)
             headingNospace = Replace(headings(1, j), " ", "_", , , vbTextCompare)
            If IsNumeric(data(i, j)) = True Then
                JS_JSONcreator = JS_JSONcreator & "             " & headingNospace & " : " & data(i, j) & "," & vbNewLine
            Else
                JS_JSONcreator = JS_JSONcreator & "             " & headingNospace & " : '" & data(i, j) & "'," & vbNewLine
            End If
        Next j
        JS_JSONcreator = JS_JSONcreator & "       " & "}," & vbNewLine
    Next i
    JS_JSONcreator = JS_JSONcreator & "];" & vbNewLine
End Function


'COUNT ARRAY DIMENTIONS
Function DimentionCounter(index As Variant) As Integer
    'This Function Counts the Columns/Dimentions in an Array
    'index is the input array
    
        On Error GoTo LC:
        For L = 1 To 100
            TempVar = index(1, L)
        Next L
LC:
        On Error GoTo 0
DimentionCounter = L -1
End Function
