'Group      Name        Start_Date      Due_Date    Comments
'String     String      Date            Date        String

Private uniqueList As Variant
Private Working As Variant
Private ArrayDims As Integer
Private ArrayRows As Single
Private YearValue As Integer
Private DisplayValuesArray As Variant
Private dateRanges As Variant
Private UniqueItemRowCount As Variant
Private dateValues As Variant
Private Datemin
Private DateMax

Public Property Let Schedule(Schedule_Array As Variant)

    Working = Schedule_Array
    Call ArrayDurations
    Call GetUniqueList
    Call SortAnArray
    Call DisplayValues
    Call AddToSheet

End Property
Private Sub AddToSheet()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim startCell As Range
    Dim i As Single
    Dim j As Single
    
    Set wb = ActiveWorkbook
    
    Application.ScreenUpdating = False
    'Add/Replace Sheet
    Application.DisplayAlerts = False
    On Error Resume Next
        Set ws = wb.Worksheets("Schedule")
        ws.Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set ws = wb.Sheets.Add
    ws.Name = "Schedule"
    
    Application.DisplayAlerts = False
    
   'Group Lables
    Set startCell = ws.Range("A3")
    Dim rA3 As Range
    Set rA3 = ws.Range("A3")
    Dim gc As Integer
    Dim g As Integer
    g = 0
    gc = 0
    
    For i = 0 To UBound(UniqueItemRowCount)
            
        If i = 0 Then
            Set startCell = ws.Range(rA3.Offset(gc, 0), rA3.Offset(gc, 0))
            gc = gc + UniqueItemRowCount(i)
            Set rng = ws.Range(startCell, startCell.Offset(UniqueItemRowCount(i) - 1, 0))
        Else
            Set startCell = ws.Range("A100000").End(xlUp).Offset(1, 0)
            Set rng = ws.Range(startCell, startCell.Offset(UniqueItemRowCount(i) - 1, 0))
        End If

        With rng
            .Value = uniqueList(i)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 90
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
            .ReadingOrder = xlContext
            .Interior.Color = RGB(88, 114, 250)
            .Font.Bold = True
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
        

        With ws.Range(rng, rng.Offset(0, UBound(dateValues) + 1))
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
        
    Next i
    
    'Date Values
    Dim pv As Variant
    pv = PivotArray(dateValues)
    
    Set rng = ws.Range("B2", Range("B2").Offset(0, UBound(dateValues)))
    With rng
        .Value = PivotArray(dateValues)
        .NumberFormat = "ddd - d mmm yy"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Interior.Color = RGB(78, 240, 180)
        .Font.Bold = True
        .Borders(xlEdgeTop).Color = vbBalck
        With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
    End With

    'Add Item to sheet
    Set startCell = ws.Range("B3")
    
    For i = 0 To UBound(DisplayValuesArray)

        Set rng = ws.Range(startCell.Offset(DisplayValuesArray(i, 8), DisplayValuesArray(i, 6)), _
                           startCell.Offset(DisplayValuesArray(i, 8), DisplayValuesArray(i, 6)))
                           
        rng.AddCommentThreaded (DisplayValuesArray(i, 4))
        
        Set rng = ws.Range(startCell.Offset(DisplayValuesArray(i, 8), DisplayValuesArray(i, 6)), _
                           startCell.Offset(DisplayValuesArray(i, 8), DisplayValuesArray(i, 7)))
                           
        With rng
            .Value = DisplayValuesArray(i, 0) & " " & DisplayValuesArray(i, 1)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
            .Interior.Color = RGB(218, 98, 150)
            .Font.Bold = True
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
    
    Next i
    ws.Range("B3").Select
    ActiveWindow.FreezePanes = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ws.Range("A1").Select
End Sub


Private Function PivotArray(toPivotArray As Variant)
    Dim i As Single
    Dim tempArray As Variant
    Dim d As Single
    d = UBound(toPivotArray)
    ReDim tempArray(0 To 0, 0 To d)
    
    For i = 0 To UBound(toPivotArray)
        tempArray(0, i) = toPivotArray(i)
    Next i
    PivotArray = tempArray
End Function

Private Sub DisplayValues()
    Dim i As Single                                                 'Loop Counter
    Dim j As Single                                                 'Loop Counter
    Dim z As Single                                                 'Loop Counter
    Dim cc As Single                                                'Item counter
    Dim ci As Single                                                'Item Group Counter
    Dim r As Integer                                                'row Counter
    Dim ir As Integer                                               'internal row Counter
    Dim c As Range                                                  'Current Cell
    Dim ds As Date                                                  'Date Start
    Dim de As Date                                                  'Date Cutoff
    Dim dl As Integer                                               'Number of days between ds and de
    Dim ids As Integer                                              'Item Start Date Range by number
    Dim ide As Integer                                              'Item End Date Range by number
    Dim uc As Integer                                               'Unique Array Item Count
    Dim rc As Variant                                               'Row Counter for Groups
    Dim rcc As Integer
    
    ReDim rc(0 To UBound(uniqueList))
    r = 0                                                           'Row Three
    rcc = 0
    Set c = cells(3, 2)                                             'Top Left Item Range
    ds = DateValue(dateRanges(0))                                   'Get Date Start
    de = DateValue(dateRanges(UBound(dateRanges)))                  'Get Date Cutoff
    dl = de - ds + 1                                                'Get Number of days between ds and de
    uc = UBound(uniqueList) + 1                                     'Unique Array Item Count
    ReDim UniqueItemRowCount(0 To uc - 1)                           'Store a count of each item against its Group

    
    'Add Date Values to Array
    ReDim dateValues(0 To dl - 1)
    For i = 0 To UBound(dateValues)                                 'Add dates to dateValues array
        dateValues(i) = ds + i
    Next i

    ReDim DisplayValuesArray(0 To UBound(Working) - 1, 0 To 9)        '9th Dim for Color | 8th is row | 7 is end date | 6 is start date
    cc = 0                                                          'Set Item Counter
    
    For i = 0 To UBound(uniqueList)                                 'Loop through each Group
        ci = 1                                                      'Set/Reset Item Group Counter
        'First Items in a Group
        

        For j = 1 To UBound(Working)                                'Loop through each item
            If Working(j, 0) = uniqueList(i) Then                   'Go through each group in order
                ir = r                                              'Reset ir | internal row counter
                ids = Working(j, 2) - ds                            'Add Item Start Date Integer for placement on a sheet
                ide = Working(j, 3) - ds                            'Add Item End Date Integer for placement on a she
                
                'Identify Conflicts
ReLook:
                For z = 0 To UBound(DisplayValuesArray)
                    If ir = DisplayValuesArray(z, 8) Then           'If item's row is = to the array item's row
                        If ids <= DisplayValuesArray(z, 7) And _
                           ide >= DisplayValuesArray(z, 6) Then      'If there is a conflict
                            ir = ir + 1                             'Add 1 to the internal Row
                            GoTo ReLook:                            'Reset Loop
                        End If
                    End If
                Next z
                
                DisplayValuesArray(cc, 0) = Working(j, 0)
                DisplayValuesArray(cc, 1) = Working(j, 1)
                DisplayValuesArray(cc, 2) = Working(j, 2)
                DisplayValuesArray(cc, 3) = Working(j, 3)
                DisplayValuesArray(cc, 4) = Working(j, 4)
                DisplayValuesArray(cc, 5) = Working(j, 5)
                DisplayValuesArray(cc, 6) = ids
                DisplayValuesArray(cc, 7) = ide
                DisplayValuesArray(cc, 8) = ir
                cc = cc + 1                                         'Item count + 1
                ci = ci + 1                                         'Add 1 to the Group item counter
            End If
            
        Next j
        
        
        
        
        
        
        For j = 0 To UBound(DisplayValuesArray)
            If DisplayValuesArray(j, 8) > r Then
                r = DisplayValuesArray(j, 8)                        'Align row with row internal
            End If
        Next j
        rcc = rcc + r
        
'###############################Not calculating teh row counter properly
        UniqueItemRowCount(i) = r
        
        r = r + 1                                                   'Go to the next row
    Next i
    
    
    For i = 1 To UBound(UniqueItemRowCount)
    
        For j = 0 To i - 1
            UniqueItemRowCount(i) = UniqueItemRowCount(i) - UniqueItemRowCount(j) '- 1
        Next j
    Next i
    UniqueItemRowCount(0) = UniqueItemRowCount(0) + 1
    
'Dim ZZZtmpAry As Variant
'ZZZtmpAry = UniqueItemRowCount
'3
'7
'6
'6
End Sub

'22 Rows
Private Sub GetUniqueList()
    Dim i As Single, items As New Collection, item
    
    items.Add (Working(1, 0))
    
    For i = 2 To UBound(Working)
        For Each item In items
            If Working(i, 0) = item Then GoTo nxt:
        Next
        
        items.Add (Working(i, 0))
nxt:
    Next i
    
    ReDim uniqueList(0 To items.Count - 1)
    For i = 0 To UBound(uniqueList)
        uniqueList(i) = items(i + 1)
    Next i
    
End Sub


Private Sub ArrayDurations()
    Dim i As Single
    Dim j As Single
    
    ArrayDims = arrayDimentionCounter(Working)
    ArrayRows = UBound(Working)
    ReDim Preserve Working(0 To ArrayRows, 0 To ArrayDims + 1)
    ArrayDims = ArrayDims + 1
    Working(0, ArrayDims) = "Duration"
    
    For i = 1 To ArrayRows
        Working(i, ArrayDims) = Int(Working(i, 3) - Working(i, 2))
    Next i
    
End Sub

Private Function arrayDimentionCounter(index As Variant) As Integer
    'This Function Counts the Columns/Dimentions in an Array
    'index is the input array
    
        On Error GoTo LC:
        For L = 1 To 100
            TempVar = index(1, L)
        Next L
LC:
        L = L - 1
        On Error GoTo 0
        arrayDimentionCounter = L
End Function

Private Sub SortAnArray()
    Dim i As Long, j As Long, k As Long
    Dim Temp As Variant
    Dim arylen As Integer
    k = 0
    
    arylen = UBound(Working) + UBound(Working) - 1
    
    ReDim dateRanges(0 To arylen)
    For j = 0 To 1
        For i = 1 To UBound(Working)
            dateRanges(k) = DateValue(Working(i, j + 2))
            k = k + 1
        Next i
    Next j
    
    'loop through bound of the arry and get the first name
    For i = LBound(dateRanges) To UBound(dateRanges) - 1
       
        'loop through again, and check if the next name is alphabetically before or after the original
        For j = i + 1 To UBound(dateRanges)
            If dateRanges(i) > dateRanges(j) Then
             
                'if the name needs to be moved before the previous name, add to a temp array
                Temp = dateRanges(j)
                
                'swop the names
                dateRanges(j) = dateRanges(i)
                dateRanges(i) = Temp
            End If
         Next j
Next i

End Sub
