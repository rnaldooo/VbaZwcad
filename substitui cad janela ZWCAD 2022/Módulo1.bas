Attribute VB_Name = "Módulo1"
Sub PegaOriginal()
Call SelectFolder(3)
End Sub

Sub PegaMirror()
Call SelectFolder(4)
End Sub

Sub PreparaOriginal()
Call ReplaceTextInFile(ActiveSheet.Range("E3").Value & ActiveSheet.Range("F3").Value, "", "")
Call ReplaceTextInFile(ActiveSheet.Range("E3").Value & ActiveSheet.Range("F3").Value, " ", "")
Call ImportTextFile(ActiveSheet.Range("E3").Value, ActiveSheet.Range("F3").Value, "TXToriginal")
Sheets("inicio").Select
End Sub



'----------------------------------------------------------------------------------------------------------------------------
'Reinaldo 25/01/2013
'
'----------------------------------------------------------------------------------------------------------------------------
Sub SelectFolder(ilinha As Integer)
    Dim diaFolder As FileDialog
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    diaFolder.Show
    ActiveSheet.Range("E" & ilinha & "").Value = diaFolder.SelectedItems(1) 'diaFolder.InitialFileName
    'ActiveSheet.Range("F" & ilinha & "").Value = Replace(diaFolder.SelectedItems(1), diaFolder.InitialFileName, "")
    Set diaFolder = Nothing
  
End Sub

Sub ImportTextFile(path As String, txtfile As String, pasta As String)
  
' setando pasta aberta
Set wb = ThisWorkbook
      
' criando guia
    Dim SheetName As String
    SheetName = pasta
    If SheetExists(SheetName) Then
    Else
       Sheets.Add.Name = SheetName
    End If
'limpando guia
    Sheets(SheetName).Select
    Cells.Select
    Selection.NumberFormat = "@"
    Selection.ClearContents
    Range("A1").Select
    
    
' criando matriz do columndatatype
    Dim i As Integer
    Dim ic As Integer
'    Dim columnDataType As XlColumnDataType
'    columnDataType = xlTextFormat
    ic = 20 ' número de coluna no arquivo
    ReDim fArray(ic) As Variant
        For i = 0 To ic
            fArray(i) = Array(i, 2)
        Next
       
    
 
'abrindo arquivo de texto
    Dim TxtFileName As String
    TxtFileName = path & txtfile
    Workbooks.OpenText _
        FileName:=TxtFileName _
        , Origin:=437 _
        , StartRow:=1 _
        , DataType:=xlDelimited _
        , TextQualifier:=xlDoubleQuote _
        , ConsecutiveDelimiter:=False _
        , Tab:=False _
        , Semicolon:=True _
        , Comma:=False _
        , Space:=False _
        , Other:=True _
        , OtherChar:="|" _
        , fieldinfo:=fArray _
        , DecimalSeparator:="," _
        , ThousandsSeparator:="." _
        , TrailingMinusNumbers:=False
                ' FieldInfo:=Array(Array(1, 2), Array(2, 2)) coluna e stilo
                ', fieldinfo:=Array(Array(1, 2), Array(2, 2), Array(3, 2))
      
' setando aquivo aberto para planilha
    Dim TMPWorkBook As Workbook
    Set TMPWorkBook = ActiveWorkbook
    Dim filesomente As String
    filesomente = Left(txtfile, InStr(txtfile, ".") - 1)
    TMPWorkBook.Sheets(filesomente).Select
    Cells.Select
    Selection.NumberFormat = "@"
' retirando os espaços
    Call ChgInfo
' copiando para planilha
    Cells.Select
    Selection.Copy
    wb.Activate
    wb.Sheets(SheetName).Select
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Cells.Select
    Cells.EntireColumn.AutoFit
    ActiveSheet.Range("A1").Select
    TMPWorkBook.Close savechanges:=False
End Sub

Function SheetExists(SheetName As String, Optional wb As Excel.Workbook)
   Dim s As Excel.Worksheet
   If wb Is Nothing Then Set wb = ThisWorkbook
   On Error Resume Next
   Set s = wb.Sheets(SheetName)
   On Error GoTo 0
   SheetExists = Not s Is Nothing
End Function

 
Sub ChgInfo()
    Dim ws              As Worksheet
    Dim Search          As String
    Dim Replacement     As String
    Dim Prompt          As String
    Dim Title           As String
    Dim MatchCase       As Boolean
    'Prompt = "What is the original value you want to replace?"
    'Title = "Search Value Input"
   ' Search = InputBox(Prompt, Title)
   ' Prompt = "What is the replacement value?"
   ' Title = "Search Value Input"
   ' Replacement = InputBox(Prompt, Title)
     Search = " "
     Replacement = ""
    For Each ws In Worksheets
        ws.Cells.Replace What:=Search, Replacement:=Replacement, LookAt:=xlPart, MatchCase:=False, ReplaceFormat:=False
        'Cells.replace What:=".", Replacement:="/", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Next
End Sub

Sub ReplaceTextInFile(SourceFile As String, _
    sText As String, rText As String)
Dim TargetFile As String, tLine As String, tString As String
Dim p As Integer, i As Long, F1 As Integer, F2 As Integer
    TargetFile = "RESULT.TMP"
    If Dir(SourceFile) = "" Then Exit Sub
    If Dir(TargetFile) <> "" Then
        On Error Resume Next
        Kill TargetFile
        On Error GoTo 0
        If Dir(TargetFile) <> "" Then
            MsgBox TargetFile & _
                " already open, close and delete / rename the file and try again.", _
                vbCritical
            Exit Sub
        End If
    End If
    F1 = FreeFile
    Open SourceFile For Input As F1
    F2 = FreeFile
    Open TargetFile For Output As F2
    i = 1 ' line counter
    Application.StatusBar = "Reading data from " & _
        TargetFile & " ..."
    While Not EOF(F1)
        If i Mod 100 = 0 Then Application.StatusBar = _
            "Reading line #" & i & " in " & _
            TargetFile & " ..."
        Line Input #F1, tLine
        If sText <> "" Then
            ReplaceTextInString tLine, sText, rText
        End If
        Print #F2, tLine
        i = i + 1
    Wend
    
    Application.StatusBar = "Closing files ..."
    Close F1
    Close F2
    Kill SourceFile ' delete original file
    Name TargetFile As SourceFile ' rename temporary file
    Application.StatusBar = False
End Sub

Private Sub ReplaceTextInString(SourceString As String, _
    SearchString As String, ReplaceString As String)
Dim p As Integer, NewString As String
    Do
        p = InStr(p + 1, UCase(SourceString), UCase(SearchString))
        If p > 0 Then ' replace SearchString with ReplaceString
            NewString = ""
            If p > 1 Then NewString = Mid(SourceString, 1, p - 1)
            NewString = NewString + ReplaceString
            NewString = NewString + Mid(SourceString, _
                p + Len(SearchString), Len(SourceString))
            p = p + Len(ReplaceString) - 1
            SourceString = NewString
        End If
        If p >= Len(NewString) Then p = 0
    Loop Until p = 0
End Sub

Sub TestReplaceTextInFile()
    ReplaceTextInFile ThisWorkbook.path & _
        "\ReplaceInTextFile.txt", "|", ";"
    ' replaces all pipe-characters (|) with semicolons (;)
End Sub


Sub LimpaPlanilhas()
    Sheets("temp").Select
    Range("A1").Select
    Sheets("temp").Select
    Columns("J:L").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Columns("R:T").Select
    Selection.ClearContents
    Columns("AL:AO").Select
    Selection.ClearContents
    Range("A1").Select
    Sheets("TXTmirror").Select
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
    Sheets("TXToriginal").Select
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
    Sheets("resultado").Select
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
    Sheets("inicio").Select
    Range("i3:i4").Select
    Selection.ClearContents
    Range("A1").Select
End Sub



Sub alimenta()

    Sheets("TXToriginal").Select
   ' Range(Selection, Selection.End(xlDown)).Select
   ' Range(Selection, Selection.End(xlToRight)).Select
    Range("A1:C7000").Select
    Selection.Copy
    Sheets("temp").Select
    Range("J3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("TXTmirror").Select
    Range("A1").Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Range(Selection, Selection.End(xlToRight)).Select
    Range("A1:C7000").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("temp").Select
    Range("R3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("inicio").Select
End Sub



Sub ordena()
Call contori
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
    
    Sheets("TXTmirror").Select
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

Sub contori()
Worksheets("inicio").Range("i3").Value = ReadNoLines_text(Worksheets("inicio").Range("E3").Value & Worksheets("inicio").Range("F3"))
Worksheets("inicio").Range("i4").Value = ReadNoLines_text(Worksheets("inicio").Range("E4").Value & Worksheets("inicio").Range("F4"))
End Sub



Function ReadNoLines_text(FileName As String) As Double

'Dimension Variables
Dim ResultStr As String
Dim FileNum As Integer
Dim CountLines As Double

'Check for no entry
If FileName = "" Then End
'Get Next Available File Handle Number
FileNum = FreeFile()
'Open Text File For Input
Open FileName For Input As #FileNum
'Set The CountLines to 1
CountLines = 1
'Loop Until the End Of File Is Reached
Do While Seek(FileNum) <= LOF(FileNum)
Line Input #FileNum, ResultStr
'Increment the CountLines By 1
CountLines = CountLines + 1
'Start Again At Top Of 'Do While' Statement
Loop
'Close The Open Text File
Close
ReadNoLines_text = CountLines - 1
End Function




Sub resultado()
    Worksheets("resultado").Range("b2").Value = "linhas"
    Worksheets("resultado").Range("b3").Value = Worksheets("inicio").Range("i3").Value
    Worksheets("resultado").Range("b4").Value = Worksheets("inicio").Range("i4").Value
    Range("C3").Select
    Sheets("temp").Select
    Range("AC1:AF1").Select
    Selection.Copy
    Sheets("resultado").Select
    Range("C3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D3").Select
    Application.CutCopyMode = False
    Range("D3").Cut Destination:=Range("C4")
    Range("F3").Select
    Selection.Cut Destination:=Range("B6")
    Range("B6").Select
    Sheets("temp").Select
    Range("AG3:AI7002").Select
    Selection.Copy
    Sheets("resultado").Select
    Range("B8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

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

Call MarcasDiferentes
Call passar2
Call arruma1

End Sub

Sub passar2()
'
'
    Sheets("temp").Select
    Range("AM3").Select
    Range("AM3:AO7002").Select
    Selection.Copy
    Sheets("resultado").Select
    Range("F8").Select
    ActiveSheet.Paste
    Range("A1").Select
 '   Columns("F:F").Select
 '   Application.CutCopyMode = False
 '   Selection.ClearContents
  '  Sheets("temp").Select
'    Columns("AL:AO").Select
 '   Selection.ClearContents
 '   Range("AN230").Select
End Sub


Sub arruma1()
'
    Sheets("resultado").Select
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(R[2]C:R[7002]C)"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("E9").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("E10").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("E9:E10").Select
    Selection.AutoFill Destination:=Range("E9:E" & Sheets("resultado").Range("a7").Value + 8 & ""), Type:=xlFillDefault
    Range("I8").Select
    ActiveCell.FormulaR1C1 = "=IF(ISNA(VLOOKUP(RC[-2],R8C3:R" & Sheets("resultado").Range("a7").Value + 8 & "C5,3,FALSE)),"""",VLOOKUP(RC[-2],R8C3:R" & Sheets("resultado").Range("a7").Value + 8 & "C5,3,FALSE))"
    Range("J8").Select
    ActiveCell.FormulaR1C1 = "=IF(ISNA(VLOOKUP(RC[-2],R8C4:R" & Sheets("resultado").Range("a7").Value + 8 & "C5,2,FALSE)),"""",VLOOKUP(RC[-2],R8C4:R" & Sheets("resultado").Range("a7").Value + 8 & "C5,2,FALSE))"
    
    Range("I8:J8").Select
    Selection.AutoFill Destination:=Range("I8:J" & Sheets("resultado").Range("f6").Value + 8 & ""), Type:=xlFillDefault
 '   Range("F8:J" & Sheets("resultado").Range("f6").Value + 8 & "").Select
 '   Range("F260").Activate
    ActiveSheet.PageSetup.PrintArea = "$F$8:$J$" & Sheets("resultado").Range("f6").Value + 8 & ""
    Range("A1").Select
End Sub


Sub Snome()
Dim SFiltro, STitulo, SArquivo As String
SFiltro = "Pastas *., *."
STitulo = "Escolha a Pasta"
SArquivo = Application.GetOpenFilename(SFiltro, , STitulo)
ActiveSheet.Range("B1").Value = SArquivo
End Sub

Function CreateFile(FileName As String, contents As String)
' creates file from string contents
 
Dim tempFile As String
Dim nextFileNum As Long
 
  nextFileNum = FreeFile
 
  tempFile = FileName
 
  Open tempFile For Output As #nextFileNum
  Print #nextFileNum, contents
  Close #nextFileNum
 
End Function

Public Function FileFolderExists(strFullPath As String) As Boolean
'Author       : Ken Puls (www.excelguru.ca)
'Macro Purpose: Check if a file or folder exists
    On Error GoTo EarlyExit
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
    
EarlyExit:
    On Error GoTo 0
End Function

Sub CRIAarq(Str1 As String, RA1 As Range, RA2 As Range)
Dim Str_arqui As String
Dim Str_cont As String
Str_arqui = RA1.Text
Str_cont = RA2.Text
If Not FileFolderExists("" & Range("B1").Text & "") Then
  MkDir (Range("B1").Text)
End If
If Not FileFolderExists("" & Range("B1").Text & "\v" & Str1 & "") Then
    MkDir (Range("B1").Text & "\v" & Str1)
End If
If RA1.Value = "lista.txt" Or RA1.Value = "lista.bat" Or RA1.Value = "criar.bat" Then
If Not FileFolderExists("" & Range("B1").Text & "\v" & Str1 & "\" & Str_arqui) Then
    CreateFile "" & Range("B1").Text & "\v" & Str1 & "\" & Str_arqui, Str_cont
End If
Else
If Not FileFolderExists("" & Range("B1").Text & "\v" & Str1 & "\" & Str_arqui & ".dtl") Then
    CreateFile "" & Range("B1").Text & "\v" & Str1 & "\" & Str_arqui & ".dtl", Str_cont
End If
End If
End Sub

Sub geraarquivos()
Dim ia As Integer
Dim Str_tem As String

For ia = 9 To 800
Str_tem = Cells(ia, 1).Text
If Str_tem <> "" Then
Call CRIAarq(Str_tem, Cells(ia, 31), Cells(ia, 32))
Call CRIAarq(Str_tem, Cells(ia, 35), Cells(ia, 36))
Else
End If

Next ia

End Sub

Function WorksheetExists(wsName As String) As Boolean
    Dim ws As Worksheet
    Dim ret As Boolean
    ret = False
    wsName = UCase(wsName)
    For Each ws In ThisWorkbook.Sheets
        If UCase(ws.Name) = wsName Then
            ret = True
            Exit For
        End If
    Next
    WorksheetExists = ret
End Function




    'Dim ColArray(,) As XlColumnDataType = {{1, XlColumnDataType.xlTextFormat}, {2, XlColumnDataType.xlTextFormat}}
    'Dim ColArray(1 To 20, 1) As XlColumnDataType
'Dim i As Integer
'For i = 1 To 20
'colarray(i,1)= {i, XlColumnDataType.xlTextFormat}
'Dim ColArray As XlColumnDataType
'For ColArray = 1 To 20
'       ColArray = 2
'Next
