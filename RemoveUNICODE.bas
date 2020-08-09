Function ClearUnwantedString(fulltext As String) As String
    Dim output As String
    Dim character As String
    For i = 1 To Len(fulltext)
        character = Mid(fulltext, i, 1)
        If (character >= "a" And character <= "z") Or (character >= "0" And character <= "9") Or (character >= "A" And character <= "Z") Then
            output = output & character
        End If
    Next
    ClearUnwantedString = output
End Function
