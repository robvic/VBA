Attribute VB_Name = "Module1"
Private Sub splitUpRegexPattern()
    Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim Myrange As Range

    Set Myrange = ActiveSheet.Range("A1:A3")

    For Each C In Myrange
        strPattern = "([^\s]+)([^\s]+)([^\s]+)"

        If strPattern <> "" Then
            strInput = C.Value
            strReplace = "$1"

            With regEx
                .Global = True
                .MultiLine = True
                .IgnoreCase = False
                .Pattern = strPattern
            End With

            If regEx.test(strInput) Then
                C.Offset(0, 1) = regEx.Replace(strInput, "$1")
                C.Offset(0, 2) = regEx.Replace(strInput, "$2")
                C.Offset(0, 3) = regEx.Replace(strInput, "$3")
            Else
                C.Offset(0, 1) = "(Not matched)"
            End If
        End If
    Next
End Sub


Function RegexExtract(ByVal text As String, _
                      ByVal extract_what As String, _
                      Optional seperator As String = "") As String

Dim i As Long, j As Long
Dim result As String
Dim allMatches As Object
Dim RE As Object
Set RE = CreateObject("vbscript.regexp")

RE.Pattern = extract_what
RE.Global = False
Set allMatches = RE.Execute(text)

For i = 0 To allMatches.Count - 1
    For j = 0 To allMatches.Item(i).SubMatches.Count - 1
        result = result & seperator & allMatches.Item(i).SubMatches.Item(j)
    Next
Next

If Len(result) <> 0 Then
    result = Right(result, Len(result) - Len(seperator))
End If

RegexExtract = result

End Function

Public Function Testthisout(number As Double) As Double
  result = number * number
  Testthisout = result
End Function

