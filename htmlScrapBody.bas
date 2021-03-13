'https://codingislove.com/parse-html-in-excel-vba/
Public HTMLScrape As Variant
Sub getPackages()
    Dim htmlBody As Variant
    Dim del As String
    Dim url As String
    
    del = ","
    url = "https://raw.githubusercontent.com/kjtakke/Excel-VBA/master/ScrapeRawWebFiles.vb"
    
    Call GetHTMLBody(del, url)
    
End Sub

Public Sub GetHTMLBody(del As String, url As String)
    Dim http As Object, html As New HTMLDocument
    Dim HTMLText As String, HTMLArray As Variant
    Dim i As Integer
    
    On Error GoTo en:
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", url, False
    http.send
    
    html.body.innerHTML = http.responseText
    HTMLText = html.body.innerHTML
    HTMLScrape = Split(HTMLText, del)
    
    GoTo fn:

en:
    msgbox ("Can not access the library server at this time")
fn:
    
End Sub
