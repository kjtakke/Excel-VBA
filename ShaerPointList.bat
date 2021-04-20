Public SP_Headers As Variant
Public SharePointLists As Collection
Public AllSharePointListData As Variant
Public ListItemCount As Integer
Public urls As Variant
Public ListIDs As Variant
Public listNames As Variant

Private Sub Class_Initialize()
    ListItemCount = 0
    ReDim urls(0 To ListItemCount)
    ReDim ListIDs(0 To ListItemCount)
    ReDim listNames(0 To ListItemCount)
    
End Sub

Public Property Let addSharePointList(url As String, listID As String, listName As String)

ReDim Preserve urls(0 To ListItemCount)
ReDim Preserve ListIDs(0 To ListItemCount)
ReDim Preserve listNames(0 To ListItemCount)
urls(ListItemCount) = url
ListIDs(ListItemCount) = listID
listNames(ListItemCount) = listName

ListItemCount = ListItemCount + 1
End Property

Sub appendArrays()
    Dim ary As Variant
    Dim k As Integer
    Dim i As Integer
    Dim j As Integer
    Dim h As Integer
    Dim c As Integer
    k = 0
    j = 1
    c = 1
    
    For Each ary In SharePointLists
        k = k + UBound(ary) + 1
    Next ary
    
    ReDim AllSharePointListData(0 To k, 0 To UBound(SP_Headers))
    'Add Headers
    For i = 0 To UBound(SP_Headers)
        AllSharePointListData(0, i) = SP_Headers(i)
    Next i
    k = 1
    
    'Append Data
    For Each ary In SharePointLists
       
        For i = 0 To UBound(ary)
            For h = 0 To UBound(SP_Headers)
            
                If IsNull(ary(i, h)) Then ary(i, h) = ""
                AllSharePointListData(j, h) = ary(i, h)
            Next h
        j = j + 1
        Next i
        c = c + 1
    Next ary
    
End Sub

Sub SharePointListsToCollection()
    Call getSharpointListHeader
    Call GetSharePointListData
    Call appendArrays
End Sub

Sub GetSharePointListData()
    Dim cnt As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim ListArray As Variant
    Dim SP_List As Variant
    Dim mySql As String
    Dim i As Integer
    Dim j As Integer
    Set SharePointLists = New Collection
    'On Error GoTo en:
    
    
    
    For j = 0 To UBound(urls)
        Set cnt = New ADODB.Connection
        Set rst = New ADODB.Recordset
        mySql = "SELECT * FROM [" & listNames(j) & "]"
    
        With cnt
            .ConnectionString = _
            "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes;DATABASE=" & urls(j) & ";LIST={" & ListIDs(j) & "};"
            .Open
        End With
    
        rst.Open mySql, cnt, adOpenDynamic, adLockOptimistic
        
        'Get Field Headers
        ReDim SP_Headers(0 To rst.Fields.Count - 1)
        
        For i = 0 To rst.Fields.Count - 1
            SP_Headers(i) = rst.Fields(i).Name
        Next i
    
        ListArray = rst.GetRows
    
        'Send SharePoint data to an Array
        ReDim SP_List(0 To UBound(ListArray, 2), 0 To UBound(ListArray)) As Variant
        Dim Counter As Long, Internalcounter As Long
        'Unpivot Array
        For Counter = 0 To UBound(ListArray)
            For Internalcounter = 0 To UBound(ListArray, 2)
                SP_List(Internalcounter, Counter) = ListArray(Counter, Internalcounter)
            Next Internalcounter
        Next Counter
        
        SharePointLists.Add SP_List
        
        If CBool(rst.State And adStateOpen) = True Then rst.Close
        Set rst = Nothing
        If CBool(cnt.State And adStateOpen) = True Then cnt.Close
        Set cnt = Nothing

    Next j
en:
End Sub


Sub getSharpointListHeader()
    Dim cnt As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim mySql As String
    Dim i As Integer
    On Error GoTo en:
    
    
    Set cnt = New ADODB.Connection
    Set rst = New ADODB.Recordset

    mySql = "SELECT * FROM [" & listNames(0) & "]"

    With cnt
        .ConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes;DATABASE=" & urls(0) & ";LIST={" & ListIDs(0) & "};"
        .Open
    End With

    rst.Open mySql, cnt, adOpenDynamic, adLockOptimistic
    
    'Get Field Headers
    ReDim SP_Headers(0 To rst.Fields.Count - 1)
    
    For i = 0 To rst.Fields.Count - 1
        SP_Headers(i) = rst.Fields(i).Name
    Next i

    If CBool(rst.State And adStateOpen) = True Then rst.Close
    Set rst = Nothing
    If CBool(cnt.State And adStateOpen) = True Then cnt.Close
    Set cnt = Nothing

en:
End Sub
