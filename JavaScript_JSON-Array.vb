'DATA
'    A     B     C
'1  Name   Qty  Age
'2  Kris   15   10
'3  Tim    30   15
'4  Jane   9    80

'INPUT
'=JS_JSONcreator("Sheet1","A2","C2","myData","A1")

'OUTPUT
'const VariableName = [
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


Function JS_JSONcreator(worksheet As String, topLeftCell As String, topRightCell As String, VariableName As String, headerTopLeftCell As String) As String
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

Function DimentionCounter(index As Variant) As Integer
    'This Function Counts the Columns/Dimentions in an Array
    'index is the input array
    
        On Error GoTo LC:
        For L = 1 To 100
            TempVar = index(1, L)
        Next L
LC:
        L = L - 1
        On Error GoTo 0
        DimentionCounter = L
End Function
