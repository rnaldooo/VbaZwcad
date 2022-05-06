Attribute VB_Name = "Módulo5"
Sub MarcasDiferentes()
Attribute MarcasDiferentes.VB_ProcData.VB_Invoke_Func = " \n14"

Sheets("temp").Select
Range("AL1:AL7002").Select
Selection.ClearContents
Range("AL1").Select

Dim idif As Integer
idif = Worksheets("temp").Range("AF1").Value

Dim ii, ia, ilinha As Integer
Dim marca1 As String

For ii = 0 To idif - 1
'Worksheets("temp").Range("AN" & ilinha + 2 & "").Value = Worksheets("temp").Range("AH" & ilinha + 2 & "").Value
marca1 = Worksheets("temp").Range("AH" & ii + 3 & "").Value
            For ia = 0 To 7000
            'Worksheets("temp").Range("AN" & ilinha + 2 & "").Value = Worksheets("temp").Range("AH" & ilinha + 2 & "").Value
            If Worksheets("temp").Range("K" & ia + 3 & "").Value = marca1 Then
            Worksheets("temp").Range("AL" & ia + 3 & "").Value = 1
            Else
            End If
            Next
Next


For ii = 0 To idif - 1
'Worksheets("temp").Range("AN" & ilinha + 2 & "").Value = Worksheets("temp").Range("AH" & ilinha + 2 & "").Value
marca1 = Worksheets("temp").Range("AI" & ii + 3 & "").Value
            For ia = 0 To 7000
            'Worksheets("temp").Range("AN" & ilinha + 2 & "").Value = Worksheets("temp").Range("AH" & ilinha + 2 & "").Value
            If Worksheets("temp").Range("M" & ia + 3 & "").Value = marca1 Then
            Worksheets("temp").Range("AL" & ia + 3 & "").Value = 1
            Else
            End If
            Next
Next


'For ii = 0 To 1000
'''Worksheets("temp").Range("AN" & ilinha + 2 & "").Value = Worksheets("temp").Range("AH" & ilinha + 2 & "").Value
'marca1 = Worksheets("temp").Range("AI" & ilinha + 2 & "").Value
        '    For ia = 0 To 7000
         ''   'Worksheets("temp").Range("AN" & ilinha + 2 & "").Value = Worksheets("temp").Range("AH" & ilinha + 2 & "").Value
'If Worksheets("temp").Range("AL" & ii + 3 & "").Value = 1 Then
         '   Worksheets("temp").Range("AL" & ilinha + 3 & "").Value = 1
         '   Else
         '   End If
         '   Next
'Next




For ia = 0 To 7000
If Worksheets("temp").Range("AL" & ia + 3 & "").Value = 1 Then
 '           If Worksheets("temp").Range("K" & ia + 3 & "").Value = marca1 Then
            Worksheets("temp").Range("AM" & ilinha + 3 & "").Value = Worksheets("temp").Range("J" & ia + 3 & "").Value
            Worksheets("temp").Range("AN" & ilinha + 3 & "").Value = Worksheets("temp").Range("K" & ia + 3 & "").Value
            Worksheets("temp").Range("AO" & ilinha + 3 & "").Value = Worksheets("temp").Range("M" & ia + 3 & "").Value
            ilinha = ilinha + 1
            Else
            End If
            Next
'
End Sub
