Private Sub Workbook_Open()
    Call worksheetAdd
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal target As Range)
    If target.Worksheet.Name <> "${agg}" Then
        On Error GoTo en:
        Call worksheetAdd
    End If
en:
End Sub

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
