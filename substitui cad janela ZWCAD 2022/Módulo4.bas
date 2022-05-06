Attribute VB_Name = "Módulo4"
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'

'
    Sheets("TXToriginal").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveWorkbook.Worksheets("TXToriginal").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("TXToriginal").Sort.SortFields.Add Key:=Range( _
        "B1:B943"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    ActiveWorkbook.Worksheets("TXToriginal").Sort.SortFields.Add Key:=Range( _
        "A1:A943"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    ActiveWorkbook.Worksheets("TXToriginal").Sort.SortFields.Add Key:=Range( _
        "C1:C943"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("TXToriginal").Sort
        .SetRange Range("A1:C943")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("inicio").Select
End Sub
Sub Macro7()
Attribute Macro7.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro7 Macro
'

'
    Range("I3:I4").Select
    Selection.Copy
    Sheets("resultado").Select
    Range(Selection, Cells(1)).Select
    Range("B3").Select
    ActiveSheet.Paste
    Range("B2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "linhas"
    Range("C3").Select
    Sheets("temp").Select
    Range("AC1:AF1").Select
    Selection.Copy
    Sheets("resultado").Select
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D3").Select
    Application.CutCopyMode = False
    Range("D3").Cut Destination:=Range("C4")
    Range("F3").Select
    Selection.Cut Destination:=Range("B6")
    Range("B6").Select
    Sheets("temp").Select
    Range("AG3:AI27").Select
    Selection.Copy
    Sheets("resultado").Select
    Range("B8").Select
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
Sub Macro8()
Attribute Macro8.VB_ProcData.VB_Invoke_Func = " \n14"

    Range("C4").Select
    Selection.Cut Destination:=Range("D6")
    Range("B4").Select
    Selection.Cut Destination:=Range("D5")
    Range("C3").Select
    Selection.Cut Destination:=Range("C6")
    Range("B3").Select
    Selection.Cut Destination:=Range("C5")
    Range("B6").Select
    Selection.Cut Destination:=Range("A7")
    Range("B7").Select
    ActiveCell.FormulaR1C1 = "id"
    Range("C7").Select
    ActiveCell.FormulaR1C1 = "marca original"
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "marca nova"
    Range("B2").Select
    Selection.Cut Destination:=Range("B6")
    Range("B6").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Cut Destination:=Range("B5")
    Range("B6").Select
    ActiveCell.FormulaR1C1 = "marcas"
    Range("B5:B6").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Cut Destination:=Range("E5:E6")
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "itens"
    Range("A7").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
