Attribute VB_Name = "Módulo10"
Sub limpararquivos()
Attribute limpararquivos.VB_ProcData.VB_Invoke_Func = " \n14"
'

'

'
    Range("E15").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    Range("E15").Select
End Sub
