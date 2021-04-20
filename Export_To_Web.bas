

'EXPORT CLASS
'This class is designed to create web dashboards in Microsoft Excel
'
'REQUIRED PACKAGES AND LIBRARIES
'   Visual Basic For Applications
'   Microsoft Access 16.0 Object Library
'   OLE Automation
'   System
'   Microsoft ActiveX Data Objects 6.1 Library
'   Microsoft ActiveX Data Objects Recordset 6.0 Library
'   Microsoft ADO Ext. 6.0 for DDL and Security
'   Microsoft Data Access Components Installed Version
'   Microsoft ADO 3.6 Object Library
'   Microsoft Outlook 16.0 Object Library
'   Microsoft Forms 2.0 Object Library (Browse for FO20.DLL)
'
'RESOURCES
'   Charts.JS
'   Bootstrap 4
'   HTML 5
'   fontsAwsome
'   jQuery
'   Allform Software Solutions
'
'REFRENCES
'
'
'
'NAMING COVENTIONS
'   Private: pv_finctionName
'   Public: finctionName
'
'   Private Variables:
'       variant: a_VariableName
'       String: s_VariableName
'       Integer: i_VariableName
'       boolean: b_VariableName
'       double: d_VariableName
'       collection: c_VariableName
'       dictionary: dict_VariableName
'       object: o_VariableName:
'           o_file
'           o_fileInstance
'           o_mail
'       Counters:
'           i, j, k, h as Single
'
'Public Subs and Functions
'    Sub HTML_Setup(rows As Integer,
'                   columns As Integer,
'                   Optional fileName As String = "",
'                   Optional filepath As String = "",
'                   Optional heading As String = "",
'                   Optional title As String = "")
'
'    Sub add_table(row As Integer,
'                  column As Integer,
'                  sql As String,
'                  Optional table_style As String = "",
'                  Optional table_class As tableClasses = 0,
'                  Optional table_id As String = "")
'
'    Sub add_metric(row As Integer,
'                   column As Integer,
'                   sql As String,
'                   Optional metric_prefix As String = "",
'                   Optional metric_sufix As String,
'                   Optional metric_style As String = "",
'                   Optional metric_class As metrics = 3,
'                   Optional metric_id As String = "")
'
'    Sub add_chart(row As Integer,
'                  column As Integer,
'                  sql As String,
'                  chart_type As chartType,
'                  chart_id As String,
'                  Optional chart_prefix As String = "",
'                  Optional chart_sufix As String = "",
'                  Optional chart_style = "",
'                  Optional chart_class As String = "",
'                  Optional chart_height As String = "400px",
'                  Optional chart_width As String = "100%",
'                  Optional chart_stacked As Boolean = False,
'                  Optional chart_legend As Boolean = True,
'                  Optional chart_colors As String = "'#00876c','#3f956d','#63a36e','#84b071','#a5bd77','#c5c980','#e5d58c','#e5bf75','#e4a862','#e39055','#e1764e','#dc5b4d','#d43d51'")
'
'    Sub add_heading(row As Integer,
'                    column As Integer,
'                    heading_tag As headings,
'                    heading_Text As String,
'                    Optional heading_style As String = "",
'                    Optional heading_class As String = "",
'                    Optional heading_id As String = "")
'
'    Sub add_div(row As Integer,
'                column As Integer,
'                div_text As String)
'
'    Sub add_styleLink(style_link As String)
'    Sub add_style(style_text As String)
'    Sub add_script_top_link(script_Link As String)
'    Sub add_scriptBottomLink(script_Link As String)
'    Sub add_scriptBottom(script_Text As String)
'
'    Sub compile()
'    Sub export(Optional loadFile As Boolean = False)
'    Sub compile_and_export(Optional loadFile As Boolean = False)
'
'    Public Function SQL_to_array(sql As String) As Variant
'
'Public Variables/Methods
'    Public HTML_Array As Variant
'    Public HTML_Column_Count As Integer
'    Public HTML_Row_Count As Integer
'    Public HTML_Script As String
'    Public HTML_Style As String
'    Public HTML_Script_Top_Links As String
'    Public HTML_Script_Bottom_Links As String
'    Public HTML_Style_Links As String
'    Public HTML_File_Name As String
'    Public HTML_File_Path As String
'    Public HTML_Elements_Count As Integer
'    Public HTML_Title As String
'    Public HTML_Heading As String
'    Public HTML_Header As Boolean
'    Public HTML_Composed As String
'
'    Public Current_Colors As Variant
'    Public Current_Icon As Variant
'    Public Current_SQL As String
'    Public Current_Array As Variant
'    Public Current_Row As Integer
'    Public Current_Column As Integer
'    Public Current_Dim_Count As Integer
'
'Enumerations
'    Public Enum tableClasses
'        Table
'        table_striped
'        table_bordered
'        table_hover
'        table_dark
'        table_dark_striped
'        table_dark_hover
'        table_borderless
'    End Enum
'
'    Public Enum chartType
'        line_chart
'        area_chart
'        bar_chart
'        hBar_chart
'        pie_chart
'    End Enum
'
'    Public Enum headings
'        h1
'        h2
'        h3
'        h4
'        h5
'        h6
'    End Enum
'
'    Public Enum metrics
'        overdue
'        due
'        completed
'        outstanding
'    End Enum
'
'EXAMPLE
'Sub TestExport()
'    Dim x As New Export_To_Web
'    Dim myArray As Variant
'
'    'Array from Sheet
'    myArray = x.RangeToArray(Worksheets("Sheet1").Range("A1:C4"))
'    x.mail = myArray
'    'Set HTML Grid
'    Call x.HTML_Setup(5, 10)
'
'    'Charts
'    Call x.add_heading(1, 1, h1, "My Dashboard Example - Proof of Concept", "text-align:center;")
'    Call x.add_chart(5, 3, myArray, bar_chart, "Chart_2", "", "", "", "", "400px", "", True)
'    Call x.add_chart(5, 2, myArray, pie_chart, "Chart_3", "", "", "", "", "400px", "", True)
'    'Tables
'
'    Call x.add_table(3, 1, myArray, "", table_striped)
'
'    'Metrics
'    myArray = x.RangeToArray(Worksheets("Sheet1").Range("B1:C4"))
'    Call x.add_metric(2, 1, myArray, "", " items", "", completed)
'    Call x.add_metric(2, 2, myArray, "", " items", "", due)
'    Call x.add_metric(2, 3, myArray, "", " items", "", outstanding)
'    Call x.add_metric(2, 4, myArray, "", " items", "", overdue)
'
'    'Compile and Export
'    Call x.compile_and_export(True)
'End Sub

'Public Constants:

Const bootstrapCSS As String = "<link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css'>"
Const bootstrapJS As String = "<script src='https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js'></script>"
Const chartsJS As String = "<script src='https://cdn.jsdelivr.net/npm/chart.js@2.8.0'></script>"
Const jQuery As String = "<script src='https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js'>"
Const fontsAwsomeCSS As String = "<link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css'>"
Const googleapis As String = "<script src='https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js'></script>"
Const cloudflare As String = "<script src='https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js'></script>"
Const dashboardCSS As String = "<link rel='stylesheet' href='https://allform-tech-200815-customer.github.io/page-templates/styles.css'>"


'Public Variables:

Public HTML_Array As Variant
Public HTML_Column_Count As Integer
Public HTML_Row_Count As Integer
Public HTML_Script As String
Public HTML_Style As String
Public HTML_Script_Top_Links As String
Public HTML_Script_Bottom_Links As String
Public HTML_Style_Links As String
Public HTML_File_Name As String
Public HTML_File_Path As String
Public HTML_Elements_Count As Integer
Public HTML_Title As String
Public HTML_Heading As String
Public HTML_Header As Boolean
Public HTML_Composed As String

Public Current_Colors As Variant
Public Current_Icon As Variant
Public Current_SQL As Variant
Public Current_Array As Variant
Public Current_Row As Integer
Public Current_Column As Integer
Public Current_Dim_Count As Integer

Public h As Single
Public i As Single
Public j As Single
Public k As Single

'Enumerations:

Public Enum tableClasses
    'table
    Table
    'table table-striped
    table_striped
    'table table-bordered
    table_bordered
    'table table-hover
    table_hover
    'table table-dark
    table_dark
    'table table-dark table-striped
    table_dark_striped
    'table table-dark table-hover
    table_dark_hover
    'table table-borderless
    table_borderless
End Enum


Public Enum chartType
    line_chart
    area_chart
    bar_chart
    hBar_chart
    pie_chart
End Enum


Public Enum headings
    h1
    h2
    h3
    h4
    h5
    h6
End Enum

Public Enum metrics
    'tile.overdue background-color: #f21313;
    overdue

    'tile.due background-color: #f08f11;
    due

    'tile.completed background-color: #000;
    completed

    'tile.open background-color: #598bff;
    outstanding

End Enum

Private Sub Class_Initialize()
    ReDim ChartHeights(0, 1)
    ChartHeightCounter = 0
End Sub


'Preview Exports
    Public Property Let preview(varMyArray As Variant)
        Dim i As Single, j As Single, k As Single
        Dim html As String
        Dim DimCount As Integer
        Dim str As String
        On Error GoTo en:

        'Get/Confirm Array Dimentions
        DimCount = arrayDimentionCounter(varMyArray)

        'HLML Tags
        html = "<!DOCTYPE html><html lang='en'><head><title>Access Table View</title><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>"
        html = html & "<link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css'>"
        html = html & "<script src='https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js'></script>"
        html = html & "<script src='https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js'></script>"
        html = html & "<script src='https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js'></script>"
        html = html & "</head>"
        html = html & "<body>"
        html = html & "<table class='table table-hover'>"
        html = html & "<thead>"

            'Cycle through Filed Headers
            For i = 0 To DimCount
                html = html & "<th>"
                html = html & varMyArray(0, i)
                html = html & "</th>"
            Next i

        html = html & "</thead>"
        html = html & "<tbody>"

            'Cycle through Filed Body
            For i = 1 To UBound(varMyArray)
                html = html & "<tr>"
                For j = 0 To DimCount
                    html = html & "<td>"

                    str = varMyArray(i, j)
                    html = html & ClearUnwantedString(str)
                    html = html & "</td>"
                Next j
                html = html & "</tr>"
            Next i
        html = html & "</tbody>"

        html = html & "</table>"
        html = html & "</body>"
        html = html & "</html>"

        'File Path to Windows Users Desktop
        filepath = "C:\Users\" & Environ("Username") & "\Desktop\table.html"
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CreateTextFile(filepath, True)

        'Write and Save HTML File
        'Debug.Print (html)
        a.WriteLine html 'ERROR
        a.Close
        Saved = True

        'Open HTML File
        Application.ActiveWorkbook.FollowHyperlink (filepath)


        'Skip Error Message
        GoTo Fn:
en:
    MsgBox ("Invalid Array")

Fn:
    End Property

    Public Property Let mail(varMyArray As Variant)
        Dim OlApp As Outlook.Application
        Dim olEmail As Outlook.MailItem
        Dim i, ii As Long
        Dim j As Single, k As Single
        Dim html As String
        Dim DimCount As Integer
        Dim str As String
        'On Error GoTo en:

        'Initilise Mail Object
        Set OlApp = New Outlook.Application
        Set olEmail = OlApp.CreateItem(olMailItem)
        'On Error GoTo en:

        'Get/Confirm Array Dimentions
        DimCount = arrayDimentionCounter(varMyArray)

        'HLML Tags
        html = "<br><br><br><!DOCTYPE html><html lang='en'><head><title>Access Table View</title><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>"
        html = html & "</head>"
        html = html & "<body>"
        html = html & "<table  cellspacing='0' border='1'>"
        html = html & "<thead>"

            'Cycle through Filed Headers
            For i = 0 To DimCount
                html = html & "<th>"
                html = html & varMyArray(0, i)
                html = html & "</th>"
            Next i

        html = html & "</thead>"
        html = html & "<tbody>"

            'Cycle through Filed Body
            For i = 1 To UBound(varMyArray)
                html = html & "<tr>"
                For j = 0 To DimCount
                    html = html & "<td>"

                    str = varMyArray(i, j)
                    html = html & ClearUnwantedString(str)
                    html = html & "</td>"
                Next j
                html = html & "</tr>"
            Next i
        html = html & "</tbody>"

        html = html & "</table>"
        html = html & "</body>"
        html = html & "</html>"

        'Create and Write/Draft Email Body Using the Users Account
        With olEmail
            .BodyFormat = olFormatHTML
            .display
            .HTMLBody = html
        End With

        'Skip Error Message
        GoTo Fn:
en:
    MsgBox ("Invalid Array")
Fn:
    End Property
  Function arrayDimentionCounter(index As Variant) As Integer
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
Function ClearUnwantedString(fulltext As String) As String
    Dim output As String
    Dim character As String
    Dim i As Single
    For i = 1 To Len(fulltext)
        
        character = Mid(fulltext, i, 1)
        
        If character = " " Or character = vbCr Or character = vbNewLine _
                           Or character = "," Or character = "." _
                           Or character = ":" Or character = "-" _
                           Or character = "/" Or character = "?" _
                           Or character = "\" Or character = "!" _
                           Or character = "$" Or character = "%" _
                           Then
        output = output & character
        GoTo en:
        End If
        
        If (character >= "a" And character <= "z") _
        Or (character >= "0" And character <= "9") _
        Or (character >= "A" And character <= "Z") Then
            output = output & character
        End If
en:
    Next
    ClearUnwantedString = output
End Function
Public Function RangeToArray(rng As Range) As Variant
    Dim tempArray As Variant
    Dim outputArray As Variant
    Dim dims As Integer
    Dim i As Integer
    Dim j As Integer
    
    tempArray = rng.Value
    dims = arrayDimentionCounter(tempArray) - 1
    rowsC = UBound(tempArray) - 1
    
    ReDim outputArray(0 To rowsC, 0 To dims)
    
    For i = 1 To UBound(tempArray)
        For j = 1 To dims + 1
            outputArray(i - 1, j - 1) = tempArray(i, j)
        Next j
    Next i
    
    RangeToArray = outputArray
    
End Function






'Public Subs:

    Sub HTML_Setup(rows As Integer, columns As Integer, Optional fileName As String = "", Optional filepath As String = "", Optional heading As String = "", Optional title As String = "")
        'This Sub sets the dimentsions for the HTML Document

        'Optional Arguments
        If Len(fileName) = 0 Then fileName = "Report"
        If Len(fileName) = 0 Then filepath = "C:\Users\" & Environ("Username") & "\Desktop\"

        'Load Public Variables
        If Len(heading) = 0 Then
            HTML_Header = False
        Else
            HTML_Header = True
        End If
        If Len(title) = 0 Then HTML_Title = "Report"
        fileName = fileName & ".html"
        filepath = filepath & fileName
        HTML_Column_Count = columns
        HTML_Row_Count = rows
        HTML_File_Name = fileName
        HTML_File_Path = filepath
        HTML_Script = ""
        HTML_Style = ""
        HTML_Elements_Count = 0
        HTML_Heading = heading
        HTML_Style = ""
        HTML_Script = ""
        HTML_Style_Links = ""
        HTML_Script_Top_Links = ""
        HTML_Script_Bottom_Links = ""
        'Set HTML_Array Dimentsions (BASE 1)
        ReDim HTML_Array(1 To rows, 1 To columns)

        'Set each array element to an empty string
        For i = 1 To rows
            For j = 1 To columns
                HTML_Array(i, j) = ""
            Next j
        Next i

    End Sub


    'Add Elements

            Sub add_table(row As Integer, column As Integer, index As Variant, Optional table_style As String = "", Optional table_class As tableClasses = 0, Optional table_id As String = "")
                'This sub creats a HTML Table from SQL->Aray

                Dim s_table_text As String
                
                Dim i_dimCount As Integer
                Current_Array = index
                Current_Dim_Count = arrayDimentionCounter(index)
                
                i_dimCount = Current_Dim_Count

                'Table Tag Optional Arguments
                s_table_text = "<table "
                If Len(table_style) > 0 Then s_table_text = s_table_text & "style='" & table_style & "' "

                If Len(table_id) > 0 Then s_table_text = s_table_text & "style='" & table_id & "' "
                s_table_text = s_table_text & vbNewLine

                If table_class = 0 Then
                    s_table_text = s_table_text & "class='table' "
                Else
                    Select Case True
                        'table
                        Case table_class = tableClasses.Table
                            s_table_text = s_table_text & "class='table' "

                        'table table-striped
                        Case table_class = tableClasses.table_striped
                            s_table_text = s_table_text & "class='table table-striped' "

                        'table table-bordered
                        Case table_class = tableClasses.table_bordered
                            s_table_text = s_table_text & "class='table table-bordered' "

                        'table table-hover
                        Case table_class = tableClasses.table_hover
                            s_table_text = s_table_text & "class='table table-hover' "

                        'table table-dark
                        Case table_class = tableClasses.table_dark
                            s_table_text = s_table_text & "class='table table-dark' "

                        'table table-dark table-striped
                        Case table_class = tableClasses.table_dark_striped
                            s_table_text = s_table_text & "class='table table-dark table-striped' "

                        'table table-dark table-hover
                        Case table_class = tableClasses.table_dark_hover
                            s_table_text = s_table_text & "class='table table-dark table-hover' "

                        'table table-borderless
                        Case table_class = tableClasses.table_borderless
                            s_table_text = s_table_text & "class='table table-borderless' "

                        Case Else
                            s_table_text = s_table_text & "class='table' "
                    End Select
                End If

                    s_table_text = s_table_text & ">" & vbNewLine

                    'Table Headers
                    s_table_text = s_table_text & "<thead>" & vbNewLine
                    s_table_text = s_table_text & "<tr>" & vbNewLine
                    For i = 0 To Current_Dim_Count
                        s_table_text = s_table_text & "<th>" & vbNewLine
                            s_table_text = s_table_text & index(0, i) & vbNewLine
                        s_table_text = s_table_text & "</th>" & vbNewLine
                    Next i
                    s_table_text = s_table_text & "</tr>" & vbNewLine
                    s_table_text = s_table_text & "</thead>" & vbNewLine

                    'Table Body
                    s_table_text = s_table_text & "<tbody>" & vbNewLine
                        For i = 1 To UBound(index)
                            s_table_text = s_table_text & "<tr>" & vbNewLine
                                For j = 0 To Current_Dim_Count
                                    s_table_text = s_table_text & "<td>" & vbNewLine
                                        s_table_text = s_table_text & index(i, j) & vbNewLine
                    'Debug.Print (index(i, j))
                                    s_table_text = s_table_text & "</td>" & vbNewLine
                                Next j
                            s_table_text = s_table_text & "</tr>" & vbNewLine
                        Next i
                    s_table_text = s_table_text & "</tbody>" & vbNewLine

                'Table close
                s_table_text = s_table_text & "</table>"

                'Load s_table_text (HTML) to HTML_Array
                HTML_Array(row, column) = HTML_Array(row, column) & s_table_text

            End Sub


            Sub add_metric(row As Integer, column As Integer, index As Variant, Optional metric_prefix As String = "", Optional metric_sufix As String, Optional metric_style As String = "", Optional metric_class As metrics = 3, Optional metric_id As String = "")
                'This Sub adds a button Metric examples at: https://allform-tech-200815-customer.github.io/page-templates/index.html

                Dim s_metric As String
                Dim s_metric_heading As String
                Dim s_metric_number As String
                Dim s_metric_class As String
                Current_Array = index
                Current_Dim_Count = arrayDimentionCounter(index)
                
                Select Case True

                    'tile.overdue background-color: #f21313;
                    Case metric_class = metrics.overdue
                        s_metric_class = "tile overdue"

                    'tile.due background-color: #f08f11;
                    Case metric_class = metrics.due
                        s_metric_class = "tile due"

                    'tile.completed background-color: #000;
                    Case metric_class = metrics.completed
                        s_metric_class = "tile completed"

                    'tile.open background-color: #598bff;
                    Case metric_class = metrics.outstanding
                        s_metric_class = "tile open"

                    Case Else
                        s_metric_class = "tile open"

                End Select

                'Assign metric elements
                s_metric_heading = index(0, 0)
                s_metric_number = index(1, 0)
                

                'Optional Arguments added to metric
                If Len(metric_prefix) > 0 Then s_metric_number = metric_prefix & s_metric_number
                If Len(metric_sufix) > 0 Then s_metric_number = s_metric_number & "<span style='font-size:50%'>" & metric_sufix & "</span>"

                s_metric = "<div align='center'>" & vbNewLine & _
                                     "<button type='button' name='button' class='" & s_metric_class & " style='" & metric_style & "' " & "id='" & metric_id & "'>" & _
                                     "<div class='tile-measure'>" & s_metric_number & "</div><br>" & vbNewLine & _
                                     "<span class='tile-comment'>" & s_metric_heading & "</span>" & vbNewLine & _
                                     "</button>" & vbNewLine & _
                                     "</div>"

                HTML_Array(row, column) = HTML_Array(row, column) & s_metric

            End Sub


            Sub add_chart(row As Integer, column As Integer, index As Variant, chart_type As chartType, chart_id As String, Optional chart_prefix As String = "", Optional chart_sufix As String = "", Optional chart_style = "", Optional chart_class As String = "", Optional chart_height As String = "400px", Optional chart_width As String = "100%", Optional chart_stacked As Boolean = False, Optional chart_legend As Boolean = True, Optional chart_colors As String = "'#00876c','#3f956d','#63a36e','#84b071','#a5bd77','#c5c980','#e5d58c','#e5bf75','#e4a862','#e39055','#e1764e','#dc5b4d','#d43d51'")

                    
                    Dim s_chart_data As String
                    Dim s_chart_lables As String
                    Dim pie_title As String
                    Dim s_chart_colors As String
                    pie_title = index(0, 0)
                    Current_Array = index
                    Current_Dim_Count = arrayDimentionCounter(index)
                    
                    'Chart Lables
                            s_chart_lables = "["
                            For i = 1 To UBound(index)
                                If i = UBound(index) Then
                                    s_chart_lables = s_chart_lables & "'" & index(i, 0) & "'"
                                Else
                                    s_chart_lables = s_chart_lables & "'" & index(i, 0) & "'" & ", "
                                End If
                            Next i
                            s_chart_lables = s_chart_lables & "]"
                    
                    'Class Canvas
                    HTML_Array(row, column) = HTML_Array(row, column) & "<div>"
                    HTML_Array(row, column) = HTML_Array(row, column) & "<canvas id='" & chart_id & "' style='height:" & chart_height & "; " & chart_width & "; " & chart_style & "' class='" & chart_class & "'></canvas>"
                    'width="400" height="400"
                    'HTML_Array(row, column) = HTML_Array(row, column) & "<canvas id='" & chart_id & "' class='" & """ b Box """ & "' height=""400"" '></canvas>"
                    
                    HTML_Array(row, column) = HTML_Array(row, column) & "</div>" & vbNewLine
                                  
                    Select Case True
                        Case chart_type = chartType.line_chart
                            s_chart_colors = chart_colors
                            HTML_Script = HTML_Script & pv_ChartScript(chart_id, "line", s_chart_lables, chart_prefix, chart_sufix, s_chart_colors, chart_stacked, chart_legend)


                        Case chart_type = chartType.pie_chart

                            'Pie Chart Data
                            s_chart_data = "["
                            For i = 1 To UBound(index)
                                If i = UBound(index) Then
                                    s_chart_data = s_chart_data & index(i, 1)
                                Else
                                    s_chart_data = s_chart_data & index(i, 1) & ", "
                                End If
                            Next i
                            s_chart_data = s_chart_data & "]"
                            
                            s_chart_colors = "[" & chart_colors & "]"
                            
                            
                            HTML_Script = HTML_Script & pv_pieChartScript(chart_id, s_chart_data, s_chart_lables, chart_prefix, chart_sufix, pie_title, s_chart_colors)
                        
                        'area
                        Case chart_type = chartType.area_chart
                        
                            s_chart_colors = chart_colors
                            HTML_Script = HTML_Script & pv_ChartScript(chart_id, "area", s_chart_lables, chart_prefix, chart_sufix, s_chart_colors, chart_stacked, chart_legend)

                        'bar
                        Case chart_type = chartType.bar_chart
                            s_chart_colors = chart_colors
                            HTML_Script = HTML_Script & pv_ChartScript(chart_id, "bar", s_chart_lables, chart_prefix, chart_sufix, s_chart_colors, chart_stacked, chart_legend)

                        'horizontalBar
                        Case chart_type = chartType.hBar_chart
                            s_chart_colors = chart_colors
                            HTML_Script = HTML_Script & pv_ChartScript(chart_id, "horizontalBar", s_chart_lables, chart_prefix, chart_sufix, s_chart_colors, chart_stacked, chart_legend)

                        Case Else


                    End Select
            End Sub


            Sub add_heading(row As Integer, column As Integer, heading_tag As headings, heading_Text As String, Optional heading_style As String = "", Optional heading_class As String = "", Optional heading_id As String = "")
                'This Sub creates a <H1-6> Tag
                Dim s_heading_text As String
                Dim s_open_Tag As String
                Dim s_close_tag As String

                'Tag
                Select Case True
                    Case heading_tag = headings.h1
                        s_open_Tag = "<h1 "
                        s_close_tag = "</h1>"
                    Case heading_tag = headings.h2
                        s_open_Tag = "<h2 "
                        s_close_tag = "</h2>"
                    Case heading_tag = headings.h3
                        s_open_Tag = "<h3 "
                        close_tag = "</h3>"
                    Case heading_tag = headings.h4
                        s_open_Tag = "<h4 "
                        s_close_tag = "</h4>"
                    Case heading_tag = headings.h5
                        s_open_Tag = "<h5 "
                        s_close_tag = "</h5>"
                    Case heading_tag = headings.h6
                        s_open_Tag = "<h6 "
                        s_close_tag = "</h6>"
                    Case Else
                        s_open_Tag = "<h1 "
                        s_close_tag = "</h1>"
                End Select

                    s_heading_text = s_open_Tag

                'Optional Arguments
                If Len(heading_style) > 0 Then
                    s_heading_text = s_heading_text & "Style='" & heading_style & "' "
                End If

                If Len(heading_class) > 0 Then
                    s_heading_text = s_heading_text & "class='" & heading_class & "' "
                End If

                If Len(heading_id) > 0 Then
                    s_heading_text = s_heading_text & "id='" & heading_id & "' "
                End If

                'Full Heading text in HTML
                s_heading_text = s_heading_text & ">" & heading_Text & s_close_tag

                'Heading HTML Text added to HTML_Array
                HTML_Array(row, column) = HTML_Array(row, column) & s_heading_text & vbNewLine

            End Sub


            Sub add_div(row As Integer, column As Integer, div_text As String)
                'This Sub is used to add custom elements to HTML_Array

                HTML_Array(row, column) = HTML_Array(row, column) & div_text

            End Sub


            'Script and Style links and code
            Sub add_styleLink(style_link As String)
                HTML_Style_Links = HTML_Style_Links & "<link rel='stylesheet' href='" & style_text & "'>" & vbNewLine
            End Sub


            Sub add_style(style_text As String)
                HTML_Style = HTML_Style & "<Style>" & vbNewLine & style_text & vbNewLine & "</style>" & vbNewLine
            End Sub


            Sub add_script_top_link(script_Link As String)
                HTML_Script_Top_Links = HTML_Script_Top_Links & "<script src='" & script_Link & "'>" & vbNewLine
            End Sub


            Sub add_scriptBottomLink(script_Link As String)
                HTML_Script_Bottom_Links = HTML_Script_Bottom_Links & "<script src='" & script_Link & "'>" & vbNewLine
            End Sub


            Sub add_scriptBottom(script_Text As String)
                HTML_Script = HTML_Script & "<script>" & vbNewLine & script_Text & vbNewLine & "</script>" & vbNewLine
            End Sub


        'Compile
                Sub export(Optional loadFile As Boolean = False)
                    
                    'File Path to Windows Users Desktop
                    filepath = HTML_File_Path
                    Set fs = CreateObject("Scripting.FileSystemObject")
                    Set a = fs.CreateTextFile("C:\Users\" & Environ("Username") & "\Desktop\" & filepath, True)
                    
                    'Write and Save HTML File
                    a.WriteLine HTML_Composed
                    a.Close
                    Saved = True
                    
                    'Open HTML File
                    Application.ActiveWorkbook.FollowHyperlink ("C:\Users\" & Environ("Username") & "\Desktop\" & filepath)
    
                End Sub


                Sub compile_and_export(Optional loadFile As Boolean = False)
                    Call compile
                    Call export(loadFile)
                End Sub
                
                
                Sub compile()
                    'This Sub compiles the HTML Document ready for export
                    
                    Dim s_html As String
                    
                    'HTML Document Open
                    s_html = "<!DOCTYPE html><html lang='en'><head><title>"
                    s_html = s_html & HTML_Title & "</title><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>"
                    
                    'Script and CSS Links
                    
                    s_html = s_html & bootstrapCSS & vbNewLine
                    s_html = s_html & bootstrapJS & vbNewLine
                    
                    s_html = s_html & chartsJS & vbNewLine
                    s_html = s_html & jQuery & vbNewLine
                    s_html = s_html & fontsAwsomeCSS & vbNewLine
                    s_html = s_html & googleapis & vbNewLine
                    s_html = s_html & cloudflare & vbNewLine
                    s_html = s_html & dashboardCSS & vbNewLine
                    
                    
                    'User Script and CSS Links
                    s_html = s_html & HTML_Style_Links & vbNewLine
                    s_html = s_html & HTML_Style & vbNewLine
                    s_html = s_html & HTML_Script_Top_Links & vbNewLine
                    
                    'HTML Body Open
                    s_html = s_html & "</head>" & vbNewLine
                    s_html = s_html & "<body>" & vbNewLine
                    
                    'Wrapper
                    s_html = s_html & "<div class='paper'>"
                    
                    'HTML_Array Input
                    For i = 1 To HTML_Row_Count
                        s_html = s_html & "<table style='width:100%; margin-left:-20px'>" & vbNewLine
                        s_html = s_html & "<tr>" & vbNewLine
                        For j = 1 To HTML_Column_Count
                            
                            If HTML_Array(i, j) <> "" Then
                                s_html = s_html & "<td style='vertical-align: top;'><div class='container-fluid'>" & vbNewLine
                                s_html = s_html & HTML_Array(i, j) & vbNewLine
                                s_html = s_html & "</div></td>" & vbNewLine
                            End If
                        Next j
                        s_html = s_html & "</tr>" & vbNewLine
                        s_html = s_html & "</table>" & vbNewLine
                    Next i
                    
                    'End Wrapper
                    s_html = s_html & "</div>"
    
                    'User Scipt Links and Scrip Tag
                    s_html = s_html & HTML_Script_Bottom_Links & vbNewLine
                    s_html = s_html & "<script>" & vbNewLine & HTML_Script & vbNewLine & "</script>" & vbNewLine
                    
                    'HTML Document Close
                    s_html = s_html & "</body>" & vbNewLine
                    s_html = s_html & "</html>" & vbNewLine
                
                    HTML_Composed = s_html
                End Sub


        'Other
                Sub to_Clipboard()

                End Sub


'Public Functions

    Public Function SQL_to_array(sql As String) As Variant
        Dim o_rst As DAO.Recordset
        Dim a_SQL As Variant
        Dim a_varField As Variant
        Dim i_dimCount As Integer
        Set o_rst = CurrentDb.OpenRecordset(sql)

        'Set Array Dimentions
        o_rst.MoveLast
        ReDim a_SQL(0 To o_rst.RecordCount, 0 To o_rst.Fields.Count - 1)
        o_rst.MoveFirst

        'Add Filed Headers To Array
        For i = 0 To o_rst.Fields.Count - 1
                a_SQL(0, i) = o_rst.Fields(i).Name
        Next i

        'SQL Body to VBA Array
        Do While Not o_rst.EOF
            For Each a_varField In o_rst.Fields
            a_SQL(o_rst.AbsolutePosition + 1, a_varField.OrdinalPosition) = a_varField
            Next a_varField
            o_rst.MoveNext
        Loop

        'Get/Confirm Array Dimentions
        Curent_Dim_Count = pv_dimentionCount(a_SQL)

        'Set Public Variables
        Current_SQL = sql
        Current_Array = a_SQL

        'Return Array
        SQL_to_array = a_SQL

    End Function


'Private Functions:

    Private Function pv_dimentionCount(index As Variant) As Integer
    'This Function Counts the Columns/Dimentions in an Array
    'index is the input array

        On Error GoTo LC:
        For i = 1 To 100
                TempVar = index(1, i)
        Next i
LC:
        i = i - 1
        On Error GoTo 0
        pv_dimentionCount = i
        Current_Dim_Count = i
    End Function



    Private Function pv_pieChartScript(pie_id, pie_data As String, pie_labels As String, Optional pie_prefix As String = "", Optional pie_sufix As String = "", Optional pie_title As String, Optional pie_colors As String) As String
        'This Function REturns the <SCRIPT> for a pie chart

        pv_pieChartScript = "var ctx = document.getElementById('" & pie_id & "').getContext('2d');" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "var myPie = new Chart(ctx, {" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "type: 'pie'," & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "data: {" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "labels: " & pie_labels & "," & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "datasets: [{" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "backgroundColor: " & pie_colors & "," & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "borderColor: '#000'," & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "borderWidth: '0px'," & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "data: " & pie_data & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "}]," & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "}," & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "options: {" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "title: {" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "display: true," & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "text: '" & pie_title & "'," & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "fontStyle: 'bold'," & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "fontSize: 20," & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "fontColor: 'black'," & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "}," & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "legend: {" & vbNewLine
        'Legend
        pv_pieChartScript = pv_pieChartScript & "position:'right'," & vbNewLine
        
        pv_pieChartScript = pv_pieChartScript & "display: true," & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "labels: {" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "fontColor: 'black'," & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "}" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "}," & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "tooltips: {" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "callbacks: {" & vbNewLine
        ' this callback is used to create the tooltip label
        pv_pieChartScript = pv_pieChartScript & "label: function(tooltipItem, data) {" & vbNewLine
        ' get the data label and data value to display
        ' convert the data value to local string so it uses a comma seperated number
        pv_pieChartScript = pv_pieChartScript & "var dataLabel = data.labels[tooltipItem.index];" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "var value = '" & pie_prefix & "' + data.datasets[tooltipItem.datasetIndex].data[tooltipItem.index].toLocaleString() + '" & pie_sufix & "';" & vbNewLine
        
        ' make this isn't a multi-line label (e.g. [["label 1 - line 1, "line 2, ], [etc...]])
        pv_pieChartScript = pv_pieChartScript & "if (Chart.helpers.isArray(dataLabel)) {" & vbNewLine
        ' show value on first line of multiline label
        ' need to clone because we are changing the value
        pv_pieChartScript = pv_pieChartScript & "dataLabel = dataLabel.slice();" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "dataLabel[0] += value;" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "} else {" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "dataLabel += value;" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "}" & vbNewLine
        
        ' return the text to display on the tooltip
        pv_pieChartScript = pv_pieChartScript & "return dataLabel;" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "}" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "}" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "}" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "}" & vbNewLine
        pv_pieChartScript = pv_pieChartScript & "});" & vbNewLine

    End Function

    Private Function chartDataSets(index As Variant, chart_colors As String, chart_type As String) As String
        Dim s_data As String, a_colors As Variant, s_bgColor As String
    
        
        'Colors Sting to Array
        a_colors = Split(chart_colors, ",")
        For i = 0 To UBound(a_colors)
            a_colors(i) = Trim(a_colors(i))
        Next i
               
        'Loop through each column
        chartDataSets = ""
        For i = 1 To Current_Dim_Count
            
            'Colors
            Select Case True
                Case chart_type = "line"
                    s_bgColor = "'rgba(255,255,255,0)'"
                Case Else
                    s_bgColor = a_colors(i - 1)
            End Select
            
            'Get data values from columns data
            s_data = "["
            For j = 1 To UBound(index)
                If j = UBound(index) Then
                    s_data = s_data & index(j, i)
                Else
                    s_data = s_data & index(j, i) & ", "
                End If
            Next j
            s_data = s_data & "]"
            
            'Build datasets
            chartDataSets = chartDataSets & "{" & vbNewLine
            chartDataSets = chartDataSets & "label:'" & index(0, i) & "'," & vbNewLine
            chartDataSets = chartDataSets & "backgroundColor: " & s_bgColor & "," & vbNewLine
            chartDataSets = chartDataSets & "borderColor: " & a_colors(i - 1) & "," & vbNewLine
            chartDataSets = chartDataSets & "borderWidth: 3," & vbNewLine
            chartDataSets = chartDataSets & "data:" & s_data & vbNewLine
            chartDataSets = chartDataSets & "}," & vbNewLine
        Next i
               
    End Function
    Private Function pv_ChartScript(chart_id As String, chart_type As String, chart_labels As String, Optional chart_prefix As String = "", Optional chart_sufix As String = "", Optional chart_colors As String, Optional chart_stacked As Boolean = False, Optional chart_legend As Boolean = True) As String
        Dim s_stacked As String, s_chart_type As String, s_user_chart_type As String
        s_user_chart_type = chart_type
        If chart_type = "area" Then s_chart_type = "line" Else s_chart_type = chart_type
        
        
        'Stacked Logic
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If chart_stacked = True Then s_stacked = "true" Else s_stacked = "false"
        
        
        'chartDataSets(Current_Array, chart_colors, chartType)
        
        pv_ChartScript = ""
        
        pv_ChartScript = pv_ChartScript & "var ctx = document.getElementById('" & chart_id & "').getContext('2d');" & vbNewLine
        pv_ChartScript = pv_ChartScript & "var chart = new Chart(ctx, {" & vbNewLine
        pv_ChartScript = pv_ChartScript & "type: '" & s_chart_type & "'," & vbNewLine

        pv_ChartScript = pv_ChartScript & "data: {" & vbNewLine
        pv_ChartScript = pv_ChartScript & "labels: " & chart_labels & "," & vbNewLine
        pv_ChartScript = pv_ChartScript & "datasets: [" & chartDataSets(Current_Array, chart_colors, s_user_chart_type) & "]" & vbNewLine
        pv_ChartScript = pv_ChartScript & "}," & vbNewLine
        pv_ChartScript = pv_ChartScript & "options: {" & vbNewLine
        pv_ChartScript = pv_ChartScript & "title: {" & vbNewLine
        pv_ChartScript = pv_ChartScript & "display: true," & vbNewLine
        pv_ChartScript = pv_ChartScript & "text: '" & Current_Array(0, 0) & "'," & vbNewLine
        pv_ChartScript = pv_ChartScript & "fontStyle: 'bold'," & vbNewLine
        pv_ChartScript = pv_ChartScript & "fontSize: 20," & vbNewLine
        pv_ChartScript = pv_ChartScript & "fontColor: 'black'," & vbNewLine
        pv_ChartScript = pv_ChartScript & "}," & vbNewLine
        pv_ChartScript = pv_ChartScript & "legend: {" & vbNewLine
        pv_ChartScript = pv_ChartScript & "position: 'right'," & vbNewLine
        pv_ChartScript = pv_ChartScript & "align: 'center'," & vbNewLine
        pv_ChartScript = pv_ChartScript & "display: true," & vbNewLine
        pv_ChartScript = pv_ChartScript & "labels: {" & vbNewLine
        pv_ChartScript = pv_ChartScript & "fontColor: 'rgb(0, 0, 0)'" & vbNewLine
        pv_ChartScript = pv_ChartScript & "}" & vbNewLine
        pv_ChartScript = pv_ChartScript & "}," & vbNewLine
        pv_ChartScript = pv_ChartScript & "scales: {" & vbNewLine
        pv_ChartScript = pv_ChartScript & "yAxes: [{" & vbNewLine
        pv_ChartScript = pv_ChartScript & "stacked: " & s_stacked & "," & vbNewLine
        pv_ChartScript = pv_ChartScript & "ticks: {" & vbNewLine
        pv_ChartScript = pv_ChartScript & "beginAtZero: true," & vbNewLine
        pv_ChartScript = pv_ChartScript & "fontColor: 'black'" & vbNewLine
        pv_ChartScript = pv_ChartScript & "}," & vbNewLine
        pv_ChartScript = pv_ChartScript & "}]," & vbNewLine
        pv_ChartScript = pv_ChartScript & "xAxes: [{"
        pv_ChartScript = pv_ChartScript & "stacked: " & s_stacked & "," & vbNewLine
        pv_ChartScript = pv_ChartScript & "ticks: {" & vbNewLine
        pv_ChartScript = pv_ChartScript & "fontColor: 'black'" & vbNewLine
        pv_ChartScript = pv_ChartScript & "}," & vbNewLine
        pv_ChartScript = pv_ChartScript & "}]" & vbNewLine
        pv_ChartScript = pv_ChartScript & "}," & vbNewLine
        pv_ChartScript = pv_ChartScript & "tooltips: {" & vbNewLine
        pv_ChartScript = pv_ChartScript & "callbacks: {" & vbNewLine
        pv_ChartScript = pv_ChartScript & "label: function(tooltipItem, data) {" & vbNewLine & vbNewLine
        
        pv_ChartScript = pv_ChartScript & "var dataLabel = data.labels[tooltipItem.index];" & vbNewLine
        pv_ChartScript = pv_ChartScript & "var value = ': " & chart_prefix & "' + data.datasets[tooltipItem.datasetIndex].data[tooltipItem.index].toLocaleString() + '" & chart_sufix & "';" & vbNewLine
        pv_ChartScript = pv_ChartScript & "if (Chart.helpers.isArray(dataLabel)) {" & vbNewLine
        pv_ChartScript = pv_ChartScript & "dataLabel = dataLabel.slice();" & vbNewLine
        pv_ChartScript = pv_ChartScript & "dataLabel[0] += value;" & vbNewLine
        pv_ChartScript = pv_ChartScript & "} else {" & vbNewLine
        pv_ChartScript = pv_ChartScript & "dataLabel += value;" & vbNewLine
        pv_ChartScript = pv_ChartScript & "}" & vbNewLine
        pv_ChartScript = pv_ChartScript & "return dataLabel;" & vbNewLine & vbNewLine
        
        pv_ChartScript = pv_ChartScript & "}" & vbNewLine
        pv_ChartScript = pv_ChartScript & "}" & vbNewLine
        pv_ChartScript = pv_ChartScript & "}" & vbNewLine
        pv_ChartScript = pv_ChartScript & "}" & vbNewLine
        pv_ChartScript = pv_ChartScript & "});" & vbNewLine
    End Function
