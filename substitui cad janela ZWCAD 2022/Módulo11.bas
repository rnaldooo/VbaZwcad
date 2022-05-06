Attribute VB_Name = "Módulo11"
Function simpleCellRegex(Myrange As Range, Mypat As String) As String
    Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String


    strPattern = Mypat ' "^[0-9]{1,3}"

    If strPattern <> "" Then
        strInput = Myrange.Value
        strReplace = ""

        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With

        If regEx.Test(strInput) Then
            simpleCellRegex = regEx.Replace(strInput, strReplace)
        Else
            simpleCellRegex = "Not matched"
        End If
    End If
End Function



Private Sub splitUpRegexPattern()
    Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim Myrange As Range

    Set Myrange = ActiveSheet.Range("A1:A3") '123a456  312A753  A1234567

    For Each c In Myrange
        strPattern = "(^[0-9]{3})([a-zA-Z])([0-9]{4})"

        If strPattern <> "" Then
            strInput = c.Value
            strReplace = "$1"

            With regEx
                .Global = True
                .MultiLine = True
                .IgnoreCase = False
                .Pattern = strPattern
            End With

            If regEx.Test(strInput) Then
                c.Offset(0, 1) = regEx.Replace(strInput, "$1")
                c.Offset(0, 2) = regEx.Replace(strInput, "$2")
                c.Offset(0, 3) = regEx.Replace(strInput, "$3")
            Else
                c.Offset(0, 1) = "(Not matched)"
            End If
        End If
    Next
End Sub
