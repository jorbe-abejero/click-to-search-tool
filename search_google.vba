Sub SearchGoogle()
    Dim personName As String
    Dim companyName As String
    Dim selectionRange As Range
    Dim lastRowE As Long, lastRowF As Long

    On Error Resume Next
    Set selectionRange = Selection
    On Error GoTo 0

    If Not selectionRange Is Nothing And selectionRange.Count = 2 Then
        ' Use highlighted cells if valid
        personName = selectionRange.Cells(1, 1).Value
        companyName = selectionRange.Cells(1, 2).Value
    Else
        ' Find the most recent (last non-empty) input in E and F
        lastRowE = ActiveSheet.Cells(Rows.Count, "E").End(xlUp).Row
        lastRowF = ActiveSheet.Cells(Rows.Count, "F").End(xlUp).Row

        ' Get the values from the last rows
        personName = ActiveSheet.Cells(lastRowE, "E").Value
        companyName = ActiveSheet.Cells(lastRowF, "F").Value

        ' Check for missing inputs first
        If Trim(personName) = "" Or Trim(companyName) = "" Then
            MsgBox "Person Name or Company Name is missing. Please check your input."
            Exit Sub
        End If

        ' Ensure both are from the same row
        If lastRowE <> lastRowF Then
            MsgBox "Person Name or Company Name is missing. Please check your input."
            Exit Sub
        End If
    End If

    ' Encode special characters for URL
    personName = EncodeSpecialChars(personName)
    companyName = EncodeSpecialChars(companyName)

    ' Open Google Search
    Dim googleURL As String
    googleURL = "https://www.google.com/search?q=" & personName & " " & companyName
    ActiveWorkbook.FollowHyperlink googleURL
End Sub
