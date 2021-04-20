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
