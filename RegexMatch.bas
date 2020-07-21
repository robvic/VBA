Attribute VB_Name = "Module1"


 Function RegExGet(aString, myExpression) As Variant
'function creates array of regular expression matches
'requires user to point vba at vbscript_regexp library
    Dim regEx As New VBScript_RegExp_55.RegExp
     Dim newArray() As String
     Dim cnt As Integer
    regEx.Pattern = myExpression
    regEx.IgnoreCase = False
    regEx.Global = True
    s = ""
    Set matches = regEx.Execute(aString)
    x = matches.Count
    ReDim newArray(x - 1) As String
    cnt = 0
        For Each Match In matches
            newArray(cnt) = Match.Value
            cnt = cnt + 1
        Next
        RegExGet = newArray()
End Function

Private Sub splitUpRegexPattern()
    Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim myRange As Range

    Set myRange = ActiveSheet.Range("A2:A5") 'Targeted values

    For Each C In myRange
        strPattern = "([0-9]{2})" 'Pattern to search

        If strPattern <> "" Then
            strInput = C.Value
            strReplace = "$1"

            With regEx
                .Global = True
                .MultiLine = True
                .IgnoreCase = True
                .Pattern = strPattern
            End With

            If regEx.Test(strInput) Then
                C.Offset(0, 1) = regEx.Replace(strInput, "$1")
            Else
                C.Offset(0, 1) = "(Not matched)"
            End If
        End If
    Next
End Sub
