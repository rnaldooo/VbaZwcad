Attribute VB_Name = "Módulo3"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Sheets("TXToriginal").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1:C943").Select
    Selection.Copy
    Sheets("Plan1 (3)").Select
    Range("J3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("TXTmirror").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1:C943").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Plan1 (3)").Select
    Range("R3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("A4:A5").Select
    Selection.AutoFill Destination:=Range("A4:A7000"), Type:=xlFillDefault
    Range("A4:A7000").Select
End Sub
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro5 Macro
'

'
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Sheets("temp").Select
    Range("J3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
