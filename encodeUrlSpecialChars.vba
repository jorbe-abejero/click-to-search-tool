' Helper function to encode special characters for URL compatibility
Function EncodeSpecialChars(inputStr As String) As String
    ' Replace special characters with their URL-encoded equivalents
    inputStr = Replace(inputStr, "&", "%26")
    inputStr = Replace(inputStr, " ", "%20")
    ' Add more replacements as needed
    EncodeSpecialChars = inputStr
End Function
