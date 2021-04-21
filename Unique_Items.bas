Option Explicit
Public TableRange As Variant, UniqueListStatus As Variant, UniqueFinalList As Variant, CounterUniqueItems As Long
Sub UniqueListStatusCaptureStatus()
    Dim CounterLoops As Long, CounterInternalUniqueLoop As Long
    Dim Unique As Boolean
    Dim TableItem As String, FirstItem As String
    CounterInternalUniqueLoop = 0
    CounterUniqueItems = 1
    TableRange = Worksheets("Data").Range("A2", Worksheets("Data").Range("F3").End(xlDown))


    Worksheets("Data").Range("J2:J132").ClearContents

For CounterLoops = 1 To UBound(TableRange, 1) - 1
TableItem = TableRange(CounterLoops, 5)
'Debug.Print (TableItem)
    If CounterLoops = 1 Then
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ReDim UniqueListStatus(10000, 0)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        UniqueListStatus(0, 0) = TableItem
        CounterUniqueItems = 2
    Else
                
            For CounterInternalUniqueLoop = 1 To CounterLoops - 1
                    'CounterInternalUniqueLoop = CounterInternalUniqueLoop + 1
                    If TableRange(CounterLoops, 5) = UniqueListStatus(CounterInternalUniqueLoop - 1, 0) Then
                        CounterInternalUniqueLoop = CounterLoops
                        Unique = False
                        'Debug.Print (TableRange(CounterLoops, 2) & "-" & UniqueListStatus(CounterInternalUniqueLoop - 1, 0))
                    Else
                        Unique = True
                        'Debug.Print (TableRange(CounterLoops, 2) & "-" & UniqueListStatus(CounterInternalUniqueLoop - 1, 0))
                    End If
    
            Next CounterInternalUniqueLoop

        If Unique = True Then
        UniqueListStatus(CounterUniqueItems - 1, 0) = TableItem
        CounterUniqueItems = CounterUniqueItems + 1
        End If
        
    End If
  
'CounterLoops = CounterLoops + 1
Next CounterLoops
TableItem = ""

Dim CountUniqueRows As Long, NextCountUniqueRow As Long

For NextCountUniqueRow = 0 To 10000

If UniqueListStatus(NextCountUniqueRow, 0) <> "" Then CountUniqueRows = CountUniqueRows + 1

Next NextCountUniqueRow
CountUniqueRows = CountUniqueRows + 1

'(Optional)


Worksheets("Data").Range("AB2", Worksheets("Data").Range("AB1").Offset(CountUniqueRows, 0)).Value = UniqueListStatus
    Worksheets("Data").Sort.SortFields.Clear
    Worksheets("Data").Sort.SortFields.Add Key:=Range("AB2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Worksheets("Data").Sort
        .SetRange Worksheets("Data").Range("AB2", Worksheets("Data").Range("AB1").Offset(CountUniqueRows, 0))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'Worksheets("Data").Range("H2").Select
UniqueFinalList = Worksheets("Data").Range("AB2", Worksheets("Data").Range("AB1").Offset(CountUniqueRows, 0)).Value


End Sub

