'This only works on data without return gaps!
'To rectify this, simpley replace in line 55:
    'Data = Worksheets("MyTmpWS").Range("A1", Worksheets("MyTmpWS").Range("A1").End(xlDown)).Value
'With:
    'Data = Worksheets("MyTmpWS").Range("A1:A10000").Value
'This will drasticlly slow down the Sub, however, you can adjust the lenght ("A1:A10000") to a smaller or large range ("A1:A100")

'Test/Example Sub
Sub TestDataFromWebToText()
    Call dataFromWebToText("https://vincentarelbundock.github.io/Rdatasets/csv/carData/Arrests.csv", "My File", ".csv")
End Sub



Sub dataFromWebToText(ByVal URL As String, fileName As String, fileExt As String)
    On Error Resume Next
    Dim str As String
    Dim ary As Variant
    ary = getDataFromWeb(URL)
    str = arrayToText(ary)
    Call writeToTextFile(fileName, fileExt, str)
End Sub

Function getDataFromWeb(URL) As Variant
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Worksheets.add().Name = "MyTmpWS"
        With Sheets("MyTmpWS").QueryTables.add(Connection:= _
            "URL;" & URL, _
            Destination:=Range("$A$1"))
            .Name = "Text"
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .WebSelectionType = xlEntirePage
            .WebFormatting = xlWebFormattingNone
            .WebPreFormattedTextToColumns = True
            .WebConsecutiveDelimitersAsOne = True
            .WebSingleBlockTextImport = False
            .WebDisableDateRecognition = False
            .WebDisableRedirections = False
            .Refresh BackgroundQuery:=False
        End With

    Dim Data As Variant
    Data = Worksheets("MyTmpWS").Range("A1", Worksheets("MyTmpWS").Range("A1").End(xlDown)).Value
    Worksheets("MyTmpWS").Delete

    Range("a1").SpecialCells (xlCellTypeLastCell)
    getDataFromWeb = Data
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Function



Function arrayToText(ary As Variant) As String
    On Error Resume Next
    arrayToText = ""
    For i = 1 To UBound(ary)
        arrayToText = arrayToText & ary(i, 1) & vbNewLine
    Next i
End Function



Sub writeToTextFile(ByVal fileName As String, fileExt As String, text As String)
    On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("C:\Users\" & Environ("userName") & "\Desktop\" & fileName & fileExt, True)
    a.WriteLine text
    a.Close
End Sub
