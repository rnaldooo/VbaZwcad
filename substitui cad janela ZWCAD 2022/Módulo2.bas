Attribute VB_Name = "Módulo2"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Sheets("TXToriginal").Select
    Cells.Select
    Selection.NumberFormat = "@"
    Selection.ClearContents
    Range("A1").Select
End Sub


Sub Replacer()
 
Dim tPath As String, tFile As String, ReplaceWhat As String, ReplaceWith As String
Dim wb As Workbook
Dim ws As Worksheet
 
'Change as required
ReplaceWhat = "Text to find"
ReplaceWith = "Text to replace"
 
'The path where your files are saved
tPath = "C:\"
 
'the *.* is all file types, *.xls will give you all xls files, *Reports.xls will give you all files ending with Reports.xls etc
tFile = Dir(tPath & "*.xls")
 
Do While Len(tFile) > 0
 
    Set wb = Workbooks.Open(tPath & tFile)
 
    'Assumes you have all data in the first sheet. Can be amended to loop through all sheets in workbook
    Set ws = wb.Sheets(1)
 
    ws.UsedRange.Replace ReplaceWhat, ReplaceWith
 
    wb.Close True
 
    tFile = Dir
 
Loop
 
 
End Sub

