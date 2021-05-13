Sub LoadListBoxValues(ctrl As Control, wb As Workbook, ws As String, rng As String, colCount As Integer)
    Application.ScreenUpdating = False
    sheets(ws).Visible = True
    wb.Worksheets(ws).Select
    With ctrl
        .ColumnCount = colCount
        If wb.Worksheets(ws).range(rng).Offset(1, 0).Value = "" Then
            .RowSource = Variables.wb.Worksheets(ws).range(rng, wb.Worksheets(ws).range(rng).Offset(1, colCount)).Address
        Else
            .RowSource = wb.Worksheets(ws).range(rng, wb.Worksheets(ws).range(rng).End(xlToRight).End(xlDown)).Address
        End If
        .ColumnHeads = True
    End With
    
    On Error Resume Next
    
    Dim j As Integer, exists As Boolean: exists = False: Dim SH1 As Worksheet
    For j = 1 To wb.sheets.Count
        If Worksheets(j).Name = "Navigation" Then exists = True
    Next j
    
    If exists = True Then
        Worksheets("Navigation").Select
    Else
        Set SH1 = Worksheets.Add
        SH1.Name = "Navigation"
        SH1.Select
    End If
    sheets(ws).Visible = False
    Application.ScreenUpdating = True
End Sub

Sub AddListBoxHeaders(headers As String, wb As Workbook, ws As String, rng As String)
    Dim headerData As Variant
    headerData = Split(headers, ",")
    Dim startRange As range:  Set startRange = wb.Worksheets(ws).range(rng)
    wb.Worksheets(ws).Cells.ClearContents
    Dim i As Integer
    For i = 0 To UBound(headerData)
        startRange.Offset(0, i).Value = headerData(i)
    Next i
End Sub

Sub loadComboBoxFromRange(ctrl As Control, wb As Workbook, ws As String, rng As String, dimention As Integer)
    Dim data As Variant: data = wb.Worksheets(ws).range(rng).CurrentRegion.Value
    Dim i As Double, str As String
    
    data = FnS_Generic.uniqueValuesFromRange(data, dimention)
    For i = 0 To UBound(data)
        ctrl.AddItem data(i)
    Next i
End Sub

Sub losdComboBoxWithSubArrayFromRange(ctrl As Control, wb As Workbook, ws As String, RngAddress As String, dimention As Integer, dimentionCompare As Integer, str As String)
    Dim data As Variant, i As Double
    data = wb.Worksheets(ws).range(RngAddress).CurrentRegion.Value

    ctrl.Clear
    For i = 1 To UBound(data)
        If data(i, dimentionCompare) = str Then ctrl.AddItem data(i, dimention)
    Next i

End Sub

Public Sub treeViewLoadValuesFromRange(ctrl As Control, wb As Workbook, ws As String, RngAddress As String)
    'This Sub loads a sheet range to a tree view control within a userform
    
    Dim i As Integer, j As Integer, k As Integer
    Dim rng As range: Set rng = wb.Worksheets(ws).range(RngAddress).CurrentRegion
    Dim treeData As Variant: treeData = rng.Value
    Dim dimCount As Integer: dimCount = FnS_Generic.arrayDimentionCounter(treeData)
    Dim newKayValue As String, parentKayValue As String, valueTreeItem As String
    
    'Load Nodes to TreeView
    For j = 2 To UBound(treeData)
        On Error Resume Next
        ctrl.Nodes.Add Key:=treeData(j, 1), text:=treeData(j, 1)
        On Error GoTo 0
    Next j
    
    'Load Children
    For i = 2 To dimCount
        For j = 2 To UBound(treeData)
            If treeData(j, i) = "" Then GoTo nxtTreeItem
            
            'Parent Key
            parentKayValue = ""
            For k = 1 To i - 1
                If k = i - 1 Then
                    parentKayValue = parentKayValue & treeData(j, k)
                Else
                    parentKayValue = parentKayValue & treeData(j, k) & "{%-%}"
                End If
            Next k
            
            'New Key
            newKayValue = ""
            For k = 1 To i
                If k = i Then
                    newKayValue = newKayValue & treeData(j, k)
                Else
                    newKayValue = newKayValue & treeData(j, k) & "{%-%}"
                End If
            Next k
            
            'Add to TreeView
            valueTreeItem = treeData(j, i)
            On Error Resume Next
            ctrl.Nodes.Add parentKayValue, tvwChild, Key:=newKayValue, text:=valueTreeItem
            On Error GoTo 0
            
nxtTreeItem:
        Next j
    Next i
    
End Sub


