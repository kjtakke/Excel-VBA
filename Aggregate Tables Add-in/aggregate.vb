Sub Main()
    Dim vals As Variant
    vals = Worksheets("${agg}").Range("A1", Worksheets("${agg}").Range("A1").End(xlDown)).Value
    If UBound(vals) < 4 Then Exit Sub
    For i = 4 To UBound(vals)
        On Error Resume Next
        Call aggregate.aggTable(vals(i, 1))
    Next i
End Sub

Sub aggTable(ByVal aggStr As String)
    Dim aggVar, args, output, tmp As Variant
    Dim fromTbl, toTbl As String
    Dim index As Long
    
    aggVar = Split(aggStr, ",")
    ReDim args(0 To UBound(aggVar) - 3)
    For i = 0 To UBound(aggVar)
        aggVar(i) = Trim(aggVar(i))
        aggVar(i) = Replace(aggVar(i), "${agg:", "")
        aggVar(i) = Replace(aggVar(i), "}", "")
    Next i
    For j = 3 To UBound(aggVar)
        args(j - 3) = Trim(aggVar(j))
    Next j
    
    Dim tmpStr As String
    tmpStr = aggVar(0)
    output = aggregate.WDTableAgg(tmpStr, UBound(args) + 1, args)
    
    Dim sn As String
    tmp = Split(aggVar(2), "!")
    
    Worksheets(tmp(0)).Range(tmp(1), Worksheets(tmp(0)).Range(tmp(1)).Offset(UBound(output) - 1, UBound(args) + 1)).Value = output
    Application.CutCopyMode = False
    On Error GoTo ev:
    
    Worksheets(tmp(0)).ListObjects.Add(xlSrcRange, Worksheets(tmp(0)).Range(tmp(1), Worksheets(tmp(0)).Range(tmp(1)).Offset(UBound(output) - 1, UBound(args) + 1)), , xlYes).Name = aggVar(1)
    
    GoTo en:
ev:
    Worksheets(tmp(0)).ListObjects(aggVar(1)).Resize Worksheets(tmp(0)).Range(tmp(1), Worksheets(tmp(0)).Range(tmp(1)).Offset(UBound(output) - 1, UBound(args) + 1))
en:
On Error GoTo 0
End Sub

Function WDTableAgg(tbl As String, cols, aggtype) As Variant
    Dim counter As Long
    Dim uWDTable, WDTable, index As Variant
    counter = o
    Agg = 0
    WDTable = Range(tbl & "[#All]").Value
    
    'ByVar error curcumvent
    Dim TempAry
    Dim ii As Long
    TempAry = WDTable
    
    Dim colsTmp As Long
    colsTmp = cols
    
    'Unique Values
    uWDTable = aggregate.WDUniqueValues(TempAry, UBound(WDTable), colsTmp)
    
    TempAry = Empty
    ReDim TempAry(1 To UBound(uWDTable), 1 To cols)
    For i = 2 To UBound(uWDTable)
        TempAry(i, 1) = uWDTable(i, 1)
    Next i
    TempAry = Empty
    
    'Fills the uWDArray with Zeros
    For i = 2 To UBound(uWDTable)
        For j = 2 To cols + 1
            uWDTable(i, j) = 0
        Next j
    Next i

    For j = 2 To cols + 1
        counter = 0
        For i = 2 To UBound(WDTable)
            For k = 1 To UBound(uWDTable)
                If WDTable(i, 1) = uWDTable(k, 1) Then
                    If aggtype(j - 2) = "sum" Or aggtype(j - 2) = "average" Then
                        If WDTable(i, j) <> "" Then
                            uWDTable(k, j) = uWDTable(k, j) + WDTable(i, j)
                            counter = counter + 1
                        End If
                    Else
                        If WDTable(i, j) <> "" Then
                            uWDTable(k, j) = uWDTable(k, j) + 1
                        End If
                    End If
                
                End If
            Next k
        Next i
        counter2 = counter2 + 1
    Next j
   
    WDTableAgg = uWDTable
End Function

Function WDUniqueValues(index As Variant, rows As Integer, cols As Long) As Variant
'Finds the unique values or the first column/dimention in an array
'index is the input array
    Dim L As Integer
    Dim C As Long
    Dim U As Boolean
    Dim UnqAryVals As Variant
    ReDim UnqAryVals(1 To UBound(index), 1 To cols + 1)
    
    'ByVar error curcumvent
    Dim TempAry
    Dim ii As Integer
    TempAry = index
        L = cols + 1
    TempAry = Empty
    For i = 1 To L
    UnqAryVals(1, i) = index(1, i)
    Next i
    
    UnqAryVals(2, 1) = index(2, 1)
    C = 2
    For i = 2 To UBound(index)
        U = True
        For j = 2 To UBound(UnqAryVals)
            If index(i, 1) = UnqAryVals(j, 1) Then U = False
        Next j
        If U = True Then
            UnqAryVals(C + 1, 1) = index(i, 1)
            C = C + 1
        End If
    Next i
    
    For j = 1 To UBound(UnqAryVals)
        If UnqAryVals(j, 1) = "" Then Exit For
    Next j
    j = j - 1
    

    ReDim TempAry(1 To j, 1 To cols + 1)
    
    For i = 1 To UBound(TempAry)
        For k = 1 To cols + 1
        TempAry(i, k) = UnqAryVals(i, k)
        Next k
    Next i
    
    WDUniqueValues = TempAry
End Function
Sub worksheetAdd()
    Application.ScreenUpdating = False
    Dim WS_Count As Integer
    Dim i As Integer
    WS_Count = ActiveWorkbook.Worksheets.Count
    For i = 1 To WS_Count
        If ActiveWorkbook.Worksheets(i).Name = "${agg}" Then GoTo en:
    Next i
    Sheets.Add.Name = "${agg}"
    With Worksheets("${agg}")
        .Range("a1").Value = "Arguments"
        .Range("A2").Value = "${agg: Input Table, Output Table, Target Location, Aggregation Type, Aggregation Type,....}"
        .Range("A3").Value = "Example:  Example:  ${agg:Table1, Table1_1, Sheet2!A1, sum, average, count}"
    End With
        Columns("A:A").EntireColumn.AutoFit
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("A2:A3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Columns("B:B").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Hidden = True
    
    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 61.5, 0.75, 88.5, 13.5). _
        Select
    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Update"
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 6). _
        ParagraphFormat
        .FirstLineIndent = 0
        .Alignment = msoAlignCenter
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 6).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 11
        .Name = "+mn-lt"
    End With
    Selection.OnAction = "Main"
    Range("A4").Select
en:
Application.ScreenUpdating = True
End Sub
