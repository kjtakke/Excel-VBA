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
