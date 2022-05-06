Attribute VB_Name = "Módulo7"
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Sheets("resultado").Select
    Range("F8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveWorkbook.Worksheets("resultado").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("resultado").Sort.SortFields.Add Key:=Range( _
        "G8:G260"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    ActiveWorkbook.Worksheets("resultado").Sort.SortFields.Add Key:=Range( _
        "F8:F260"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    ActiveWorkbook.Worksheets("resultado").Sort.SortFields.Add Key:=Range( _
        "H8:H260"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("resultado").Sort
        .SetRange Range("F8:H260")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=6
End Sub
Sub arruma1()
Attribute arruma1.VB_ProcData.VB_Invoke_Func = " \n14"
'
    Sheets("resultado").Select
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("E9").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("E10").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("E9:E10").Select
    Selection.AutoFill Destination:=Range("E9:E" & Sheets("resultado").Range("a7").Value + 8 & ""), Type:=xlFillDefault
    Range("I8").Select
    ActiveCell.FormulaR1C1 = "=IF(ISNA(VLOOKUP(RC[-2],R8C3:R26C5,3,FALSE)),"""",VLOOKUP(RC[-2],R8C3:R26C5,3,FALSE))"
    Range("J8").Select
    ActiveCell.FormulaR1C1 = "=IF(ISNA(VLOOKUP(RC[-2],R8C4:R26C5,2,FALSE)),"""",VLOOKUP(RC[-2],R8C4:R26C5,2,FALSE))"
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(R[2]C:R[7002]C)"
    Range("I8:J8").Select
    Selection.AutoFill Destination:=Range("I8:J" & Sheets("resultado").Range("f6").Value + 8 & ""), Type:=xlFillDefault
 '   Range("F8:J" & Sheets("resultado").Range("f6").Value + 8 & "").Select
 '   Range("F260").Activate
    ActiveSheet.PageSetup.PrintArea = "$F$8:$J$" & Sheets("resultado").Range("f6").Value + 8 & ""
    Range("A1").Select
End Sub
Sub Macro10()
Attribute Macro10.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro10 Macro

    Dim ColCnt As Integer
    Dim rng As Range
    Dim cw As String
    Dim c As Integer
    
    ColCnt = ActiveSheet.UsedRange.Columns.Count
    Set rng = ActiveSheet.UsedRange
    With ListBox1
        .ColumnCount = ColCnt
        .RowSource = rng.Address
        cw = ""
        For c = 1 To .ColumnCount
            cw = cw & rng.Columns(c).Width & ";"
        Next c
        .ColumnWidths = cw
        .ListIndex = 0
    End With
'

End Sub
