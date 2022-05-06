Attribute VB_Name = "Módulo8"
Option Explicit

'crialayer ok
'dmargem ok


'===========================================================================================
'==== crialayer ================================================ Reinaldo === 01/04/2009 ===
'
'cria os layers a partir da planilha
'adicione as referências:
'1- autocad 2010 type library (ou a sua versão)
'2- autocad/objectdbx commom 18 type libray (ou a sua versão do cad)

Dim ZWCAD As ZcadApplication '                                             Variáveis Globais
Dim ad As ZcadDocument '                                                   Variáveis Globais
Dim WExcel As Excel.Sheets '                                               Variáveis Globais

Sub null_cria_layer() '                                                     Função principal
    Dim WExcel As Excel.Worksheet
    Dim AL_layer As ZWcadLayer
    Dim i(0 To 1) As Integer
    Set WExcel = Application.Worksheets("inicio")
    If B_pega_desenho() = True Then
    Else
        MsgBox ("Erro: Abra o ZWcad!" & Err.Description)
        Err.Clear
        Exit Sub
    End If
  
    On Error Resume Next
    ad.Linetypes.Load "HIDDEN", "ZWcad.lin" '                       carregando tipo de linhas
    ad.Linetypes.Load "CENTER", "ZWcad.lin" '                       carregando tipo de linhas
    Sheets("cria layer").Select '                                   le valores em cria layer
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(RC[-12]:R[100]C[-12])"
    i(1) = Range("M1").Value

    For i(0) = 1 To i(1)
        If WExcel.Cells(i(0), 1) <> "" Then
            Set AL_layer = AL_verifica_ou_cria_layer(WExcel.Cells(i(0), 1))
            AL_layer.Color = WExcel.Cells(i(0), 2)
            AL_layer.LineType = CStr(WExcel.Cells(i(0), 3))
        End If
    Next i(0)
    Sheets("verificarSec").Select
    MsgBox ("Pronto!!" & Chr(10) & "Layers criados")
End Sub


Function B_pega_desenho() As Boolean            '                                 pega o desenho ativo
    On Error GoTo erro
    Dim zwApp As ZWcadApplication
    Dim ad As ZWcadDocument
    'acerte para a versão do autocad 2010 => 18  2009 => 17.2  2008 => 17.1  2007 => 17.0
    'Application.Worksheets("inicio").Range("H11").Value
    Set zwApp = GetObject(, Application.Worksheets("inicio").Range("H11").Value) 'Application.Worksheets("inicio").Range("H11").Value)
    Set ad = zwApp.ActiveDocument
ok:
    B_pega_desenho = True
    Exit Function
erro:
    B_pega_desenho = False
End Function

Function AL_verifica_ou_cria_layer(ByVal nome As String) As ZcadLayer 'verf ou cria layer
    On Error GoTo cria
    Set AL_verifica_ou_cria_layer = ad.Layers.Item(nome)
    Exit Function
cria:
    Set AL_verifica_ou_cria_layer = ad.Layers.Add(nome)
End Function



'--------------------------------------------------------------------------------------------------------------------------------
Sub dmargem()   ' reinaldo - 23/04/2009  - fazer cálculos para dimensionamento de seção de concreto
'--------------------------------------------------------------------------------------------------------------------------------

Dim zwApp As ZWcadApplication
Dim zwdoc As ZWcadDocument

' AUTOCAD**** - definindo desenho aberto
Set zwApp = GetObject(, Application.Worksheets("inicio").Range("H11").Value)
Dim aAD As ZWCAD.ZWcadDocument: Set aAD = ZWCAD.ActiveDocument:      aAD.Activate   ' aAD - ZWcad document
Dim aMS As ZWcadModelSpace:       Set aMS = ZWCAD.ActiveDocument.ModelSpace           ' aMS - ZWcad modelspace

Dim plineObj As ZWcadLWPolyline
Dim points(0 To 9) As Double
    
Dim Mee As Excel.Worksheet
Set Mee = Application.Worksheets("inicio")

Dim vx, vy As Variant

        vx = Mee.Range("I1").Value
        vy = Mee.Range("J1").Value
    
    ' Define the 2D polyline points
    points(0) = 0: points(1) = 0
    points(2) = vx: points(3) = 0
    points(4) = vx: points(5) = vy
    points(6) = 0: points(7) = vy
    points(8) = 0: points(9) = 0
        
    Set plineObj = ZWCAD.ActiveDocument.PaperSpace.AddLightWeightPolyline(points)
    'ZWcad.ActiveDocument.Regen
'    ZWcad.ActiveDocument.ActiveViewport.ZoomAll

        
    ' Create a lightweight Polyline object in model space
    Set plineObj = aAD.PaperSpace.AddLightWeightPolyline(points)
    plineObj.Layer = "0"
    
    
    points(0) = 1.5: points(1) = 1
    points(2) = vx - 1: points(3) = 1
    points(4) = vx - 1: points(5) = vy - 1
    points(6) = 1.5: points(7) = vy - 1
    points(8) = 1.5: points(9) = 1
        
    ' Create a lightweight Polyline object in model space
    Set plineObj = aAD.PaperSpace.AddLightWeightPolyline(points)
    plineObj.Layer = "0"
    ZWCAD.ActiveDocument.Regen


End Sub







Sub arrumartud()
Dim varUserInput As Variant
Dim objZCAD As ZcadApplication
Dim objDOC As ZcadDocument
Dim objNEWSS As ZcadSelectionSet
Dim varPT1 As Variant
Dim intGroupCode(0 To 7) As Integer ' caso com dim 0 to 4
Dim varGroupValue(0 To 7) As Variant ' caso com dim 0 to 4
'Dim entTypeConstant As ZWcadBlockReference ' String
Dim i As Integer
Dim attribs As Variant
Dim stexto, svalor, ssubsti, stemp, ssnome As String
Dim iss, idwg, ifdwg, itt, itemtexto, icontem, nnlinha, ivv As Integer
Dim Mee As Excel.Worksheet
Dim ba As Boolean
Dim dposx, dposy As Double
Dim dinpoint As ZcadPoint
Dim vaaa As Variant

Worksheets("resultado").Select
Cells.Select
Selection.ClearContents
Range("A1").Select
Worksheets("inicio").Select

ivv = -1
ivv = ivv + 1: intGroupCode(ivv) = -4: varGroupValue(ivv) = "<NOT"
ivv = ivv + 1: intGroupCode(ivv) = 0: varGroupValue(ivv) = "DIMENSION"
ivv = ivv + 1: intGroupCode(ivv) = -4: varGroupValue(ivv) = "NOT>"  ' filtro dimensao
ivv = ivv + 1: intGroupCode(ivv) = -4: varGroupValue(ivv) = "<OR"
ivv = ivv + 1: intGroupCode(ivv) = 0: varGroupValue(ivv) = "insert"
ivv = ivv + 1: intGroupCode(ivv) = 0: varGroupValue(ivv) = "text"
ivv = ivv + 1: intGroupCode(ivv) = 0: varGroupValue(ivv) = "mtext"
ivv = ivv + 1: intGroupCode(ivv) = "-4": varGroupValue(ivv) = "OR>"

'If OZWcadEstaAberto() Then
'Else
Call ZWcadInstance
'End If

Set objZCAD = GetObject(, "" & Application.Worksheets("inicio").Range("H11").Value & "")
ifdwg = Sheets("inicio").Range("e14").Value
ba = Sheets("lista").Range("b3").Value

        For idwg = 1 To ifdwg                                                   ' para cada arquivo
        
                'Set objDOC = objZCAD.ActiveDocument
                Set objDOC = objZCAD.Documents.Open("" & Sheets("inicio").Range("e" & idwg + 14 & "") & "")

                On Error Resume Next
                objDOC.SelectionSets.Item("OWITBL").Delete
                Err.Clear
                Set objNEWSS = objDOC.SelectionSets.Add("OWITBL")
                '----- Form1.Hide
                '----- PickOnScreen:
                If Sheets("inicio").Range("e11").Value = "Janela" Then
                
                    Dim Vponto_1(0 To 2) As Double
                    Dim Vponto_2(0 To 2) As Double
                    
                    Vponto_1(0) = Sheets("inicio").Range("g7").Value
                    Vponto_1(1) = Sheets("inicio").Range("h7").Value
                    Vponto_1(2) = 0
                    
                    Vponto_2(0) = Sheets("inicio").Range("g8").Value
                    Vponto_2(1) = Sheets("inicio").Range("h8").Value
                    Vponto_2(2) = 0
                    
                    
                    '-----  object.Select(Type, [Point1], [Point2], [FilterType], [FilterData])
                    
                    objNEWSS.Select zcSelectionSetWindow, Vponto_1, Vponto_2, intGroupCode, varGroupValue
                    'objNEWSS.Select (acSelectionSetWindow, Vponto_1, Vponto_2, intGroupCode, varGroupValue)
                Else
                    objNEWSS.Select zcSelectionSetAll, , , intGroupCode, varGroupValue
                End If
                
                ' ----- objNEWSS.Select.all intGroupCode, varGroupValue
                ' ------ objNEWSS.SelectOnScreen intGroupCode, varGroupValue
                Set Mee = Application.Worksheets("lista")
               
               Worksheets("resultado").Select
               Range("A1").Select
               Sheets("resultado").Cells(nnlinha + 1, 2).Value = objZCAD.ActiveDocument.FullName
               Cells(nnlinha + 1, 2).Select
               
                Select Case Sheets("inicio").Range("e9").Value
                Case "substituir"
                            For iss = 1 To Mee.Range("b2").Value                ' para cada substituição
                                svalor = Mee.Range("a" & iss + 4 & "").Value
                                ssubsti = Mee.Range("b" & iss + 4 & "").Value
                                    For i = 0 To objNEWSS.Count - 1             ' para cada texto
                                        attribs = objNEWSS.Item(i)
                                        stexto = objNEWSS.Item(i).TextString
                                        If ba Then ' no texto inteiro
                                             If stexto = svalor Then
                                                 'objNEWSS.Item(i).TextString = ssubsti
                                                 'Sheets("resultado").Cells(nnlinha + 1, 2).Value = objZCAD.ActiveDocument.FullName
                                                 'Sheets("resultado").Cells(nnlinha + 1, 3).Value = svalor
                                                 'Sheets("resultado").Cells(nnlinha + 1, 4).Value = ssubsti
                                                 'Sheets("resultado").Cells(nnlinha + 1, 5).Value = stexto
                                                ' Sheets("resultado").Cells(nnlinha + 1, 4).Select
                                                 'nnlinha = nnlinha + 1
                                              Else
                                              End If
                                        Else ' em cada parte do texto
                                            itemtexto = 0
                                            icontem = 0
                                            itemtexto = InStr(1, stexto, svalor, vbTextCompare)
                                            ssnome = objNEWSS.Item(i).ObjectName
                                            For itt = 1 To Mee.Range("e2").Value
                                                icontem = InStr(1, stexto, Mee.Range("e" & itt + 4 & "").Value, vbTextCompare) + icontem
                                            Next
                                             
                                            If itemtexto > 0 Then 'tem o texto
                                                      If icontem = 0 Then  'vendo de se contem os textos     stexto <> stemp Then
                                                         stemp = ReplaceTextInString("" & stexto & "", "" & svalor & "", "" & ssubsti & "")
                                                         objNEWSS.Item(i).TextString = stemp
                                                         Sheets("resultado").Cells(nnlinha + 1, 2).Value = objZCAD.ActiveDocument.FullName
                                                         Sheets("resultado").Cells(nnlinha + 1, 3).Value = svalor
                                                         Sheets("resultado").Cells(nnlinha + 1, 4).Value = ssubsti
                                                         Sheets("resultado").Cells(nnlinha + 1, 5).Value = stexto
                                                         Sheets("resultado").Cells(nnlinha + 1, 6).Value = stemp
                                                         Sheets("resultado").Cells(nnlinha + 1, 7).Value = "" 'objNEWSS.Item(i).ObjectID
                                                         Sheets("resultado").Cells(nnlinha + 1, 8).Value = ssnome
                                                         Sheets("resultado").Cells(nnlinha + 1, 9).Value = objNEWSS.Item(i).Handle
                                                         Sheets("resultado").Cells(nnlinha + 1, 10).Value = objNEWSS.Item(i).Layer
                                                         Sheets("resultado").Cells(nnlinha + 1, 11).Value = objNEWSS.Item(i).ObjectID32
                                                         Sheets("resultado").Cells(nnlinha + 1, 12).Value = objNEWSS.Item(i).InsertionPoint(1)
                                                         Sheets("resultado").Cells(nnlinha + 1, 13).Value = objNEWSS.Item(i).InsertionPoint(2)
                                                         Sheets("resultado").Cells(nnlinha + 1, 14).Value = objNEWSS.Item(i).InsertionPoint(3)
                                                         Sheets("resultado").Cells(nnlinha + 1, 15).Value = objNEWSS.Item(i).StyleName
                                                         stemp = ""
                                                         nnlinha = nnlinha + 1
                                                      Else
                                                      End If
                                              Else
                                              End If
                                        End If
                                    Next    ' para cada texto
                                   ' ZWcad.ActiveDocument.Regen (True)
                            Next            ' para cada substituição
                                      ' para cada arquivo
            
            
            Case "trocar"
                           For i = 0 To objNEWSS.Count - 1            ' para cada texto
                                        attribs = objNEWSS.Item(i)
                                        stexto = objNEWSS.Item(i).TextString
                                        Sheets("resultado").Cells(nnlinha + 1, 2).Value = objZCAD.ActiveDocument.FullName
                                        Sheets("resultado").Cells(nnlinha + 1, 3).Value = stexto
                                        stexto = Right(stexto, Len(stexto) - 1)
                                        stexto = Sheets("lista").Range("h3").Value & stexto
                                        Sheets("resultado").Cells(nnlinha + 1, 4).Value = stexto
                                        Sheets("resultado").Cells(nnlinha + 1, 4).Select
                                        objNEWSS.Item(i).TextString = stexto
                                        'nnlinha = nnlinha + 1
                            Next
            
             Case "adicionar"
                           For i = 0 To objNEWSS.Count - 1            ' para cada texto
                                        attribs = objNEWSS.Item(i)
                                        stexto = objNEWSS.Item(i).TextString
                                        Sheets("resultado").Cells(nnlinha + 1, 2).Value = objZCAD.ActiveDocument.FullName
                                        Sheets("resultado").Cells(nnlinha + 1, 3).Value = stexto
                                        'stexto = Right(stexto, Len(stexto) - 1)
                                        stexto = stexto & Sheets("lista").Range("h3").Value
                                        Sheets("resultado").Cells(nnlinha + 1, 4).Value = stexto
                                        Sheets("resultado").Cells(nnlinha + 1, 4).Select
                                        objNEWSS.Item(i).TextString = stexto
                                        'nnlinha = nnlinha + 1
                                    Next
            
            
                Case "multiplicar"
                                    For i = 0 To objNEWSS.Count - 1            ' para cada texto
                                        attribs = objNEWSS.Item(i)
                                        stexto = objNEWSS.Item(i).TextString
                                        stexto = Replace(stexto, ".", ",")
                                        Sheets("resultado").Cells(nnlinha + 1, 2).Value = objZCAD.ActiveDocument.FullName
                                        Sheets("resultado").Cells(nnlinha + 1, 3).Value = stexto
                                        Dim nnvalor As Double
                                        nnvalor = CDbl(stexto)
                                        nnvalor = nnvalor * Sheets("lista").Range("h2").Value
                                        stexto = CStr(nnvalor)
                                        stexto = Replace(stexto, ",", ".")
                                        Sheets("resultado").Cells(nnlinha + 1, 4).Value = stexto
                                        Sheets("resultado").Cells(nnlinha + 1, 4).Select
                                        objNEWSS.Item(i).TextString = stexto
                                        nnlinha = nnlinha + 1
                                    Next


                Case "listar"
                                    For i = 0 To objNEWSS.Count - 1            ' para cada texto
                                        
                                        attribs = objNEWSS.Item(i)
                                        stexto = objNEWSS.Item(i).TextString
                                        stexto = Replace(stexto, ".", ",")
                                        'Sheets("resultado").Cells(nnlinha + 1, 2).Value = objZCAD.ActiveDocument.FullName
                                        Sheets("resultado").Cells(nnlinha + 1, 3).Value = stexto
                                        Sheets("resultado").Cells(nnlinha + 1, 4).Select
                                        'nnlinha = nnlinha + 1
                                    Next
                Case "deslocar"
                            'For iss = 1 To Mee.Range("b2").Value                ' para cada substituição
                            'svalor = Mee.Range("a" & iss + 4 & "").Value
                           ' ssubsti = Mee.Range("b" & iss + 4 & "").Value
                                    For i = 0 To objNEWSS.Count - 1             ' para cada texto
                                        attribs = objNEWSS.Item(i)
                                        stexto = objNEWSS.Item(i).TextString
                                        If ba Then ' no texto inteiro
                                        'If stexto = svalor Then
                                        
                                        vaaa = objNEWSS.Item(i).InsertionPoint
                                        objNEWSS.Item(i).Alignment = zcAlignmentRight
                                        vaaa(0) = vaaa(0) + Sheets("lista").Range("k2").Value
                                        vaaa(1) = vaaa(1) + Sheets("lista").Range("k3").Value
                                          
                                        objNEWSS.Item(i).InsertionPoint = vaaa
                                        
                                        Sheets("resultado").Cells(nnlinha + 1, 2).Value = objZCAD.ActiveDocument.FullName
                                        Sheets("resultado").Cells(nnlinha + 1, 3).Value = stexto
                                        Sheets("resultado").Cells(nnlinha + 1, 4).Value = vaaa(0)
                                        Sheets("resultado").Cells(nnlinha + 1, 5).Value = vaaa(1)
                                        Sheets("resultado").Cells(nnlinha + 1, 4).Select
                                        nnlinha = nnlinha + 1
                                                                               
                                        'Else
                                        'End If
                                        Else ' em cada parte do texto
                                        
                                        End If
                                    Next    ' para cada texto
                            'Next            ' para cada substituição
                                      ' para cada arquivo
                
                Case "pegarcoord"
                       
                           For iss = 1 To Mee.Range("b2").Value                ' para cada substituição
                            svalor = Mee.Range("a" & iss + 4 & "").Value
                            ssubsti = Mee.Range("b" & iss + 4 & "").Value
                            Sheets("resultado").Cells(nnlinha + 1, 2).Value = objZCAD.ActiveDocument.FullName
                            
                                    For i = 0 To objNEWSS.Count - 1             ' para cada texto
                                        attribs = objNEWSS.Item(i)
                                        stexto = objNEWSS.Item(i).TextString
                                        If ba Then ' no texto inteiro
                                        If stexto = svalor Then
                                        vaaa = objNEWSS.Item(i).InsertionPoint
                                        
                                        
                                        Dim objTTT As ZcadSelectionSet
                                        On Error Resume Next
                                        objDOC.SelectionSets.Item("ATTTTT").Delete
                                        Err.Clear
                                        Set objTTT = objDOC.SelectionSets.Add("ATTTTT")
                                        
                                        If Sheets("lista").Range("K5").Value = "TRUE" Then
                
                                        Dim Vponto_a1(0 To 2) As Double
                                        Dim Vponto_a2(0 To 2) As Double
                                        
                                        Vponto_a1(0) = vaaa(0) + Sheets("lista").Range("k7").Value
                                        Vponto_a1(1) = vaaa(1) + Sheets("lista").Range("k87").Value
                                        Vponto_a1(2) = 0
                                        
                                        Vponto_a2(0) = vaaa(0) + Sheets("lista").Range("k7").Value + Sheets("lista").Range("k10").Value
                                        Vponto_a2(1) = vaaa(1) + Sheets("lista").Range("k87").Value + Sheets("lista").Range("k11").Value
                                        Vponto_a2(2) = 0
                    
                    
                                            '-----  object.Select(Type, [Point1], [Point2], [FilterType], [FilterData])
                                            
                                            objTTT.Select zcSelectionSetWindow, Vponto_a1, Vponto_a2, intGroupCode, varGroupValue
                                        Else
                                          
                                        End If
                                        
                                        
                                        If objTTT.Count = 1 Then
                                        ssubsti = objTTT.Item(0).TextString
                                        vaaa = objTTT.Item(0).InsertionPoint
                                        Else
                                        End If
                                        
                                        'Sheets("resultado").Cells(nnlinha + 1, 2).Value = objZCAD.ActiveDocument.FullName
                                        Sheets("resultado").Cells(nnlinha + 1, 3).Value = svalor
                                        Sheets("resultado").Cells(nnlinha + 1, 4).Value = ssubsti
                                        'Sheets("resultado").Cells(nnlinha + 1, 5).Value = vaaa(0)
                                        'Sheets("resultado").Cells(nnlinha + 1, 6).Value = vaaa(1)
                                        Sheets("resultado").Cells(nnlinha + 1, 4).Select
                                        
                                        
                                        Else
                                        End If
                                        Else ' em cada parte do texto
                                        stemp = ReplaceTextInString("" & stexto & "", "" & svalor & "", "" & ssubsti & "")
                                        vaaa = objNEWSS.Item(i).InsertionPoint
                                        objNEWSS.Item(i).TextString = stemp
                                        stemp = ""
                                        'Sheets("resultado").Cells(nnlinha + 1, 2).Value = objZCAD.ActiveDocument.FullName
                                        Sheets("resultado").Cells(nnlinha + 1, 3).Value = svalor
                                        Sheets("resultado").Cells(nnlinha + 1, 4).Value = ssubsti
                                        Sheets("resultado").Cells(nnlinha + 1, 5).Value = vaaa(0)
                                        Sheets("resultado").Cells(nnlinha + 1, 6).Value = vaaa(1)
                                        Sheets("resultado").Cells(nnlinha + 1, 4).Select
                                        
                                        End If
                                    Next    ' para cada texto
                                    'nnlinha = nnlinha + 1
                                    'ZWcad.ActiveDocument.Regen (True)
                            Next            ' para cada substituição
                                      ' para cada arquivo
                               
                               
                
                Case Else
                End Select
       ' nnlinha = nnlinha + 1
        'aAD.Regen (True)
        'ZWcad.ActiveDocument.Regen (True)
        objDOC.Regen (True)
        'objDOC.Regen
        objDOC.Save
        objDOC.Close
             
        Next
'ZWcad.ActiveDocument.Regen
If Not objNEWSS Is Nothing Then objNEWSS.Delete
'Form1.Show
'objZCAD.silentOperation = True
objZCAD.ActiveDocument.Close (False)
'objZCAD.DisplayAlerts = False
objZCAD.Quit
'objZCAD.DisplayAlerts = True
'objZCAD.silentOperation = False
'objDOC.Quit
'ZWcad.Quit
End Sub

Private Function ReplaceTextInString(SourceString As String, SearchString As String, ReplaceString As String)
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
    ReplaceTextInString = SourceString
End Function


Function OExcelEstaAberto() As Boolean
Dim xlApp As Excel.Application
On Error Resume Next
Set xlApp = GetObject(, "Excel.Application")
OExcelEstaAberto = (Err.Number = 0)
Set xlApp = Nothing
Err.Clear
End Function

Sub ExcelInstance()
Dim eAp As Excel.Application

Dim ExcelAberto As Boolean

ExcelAberto = OExcelEstaAberto()
If ExcelAberto Then
Set eAp = GetObject(, "Excel.Application")
Else
Set eAp = CreateObject("Excel.Application")
End If
eAp.Visible = True
'If Not ExcelAberto Then eAp.Quit
Set eAp = Nothing
End Sub

Function OZWcadEstaAberto() As Boolean
Dim zwApp As ZcadApplication
On Error Resume Next
Set zwApp = GetObject(, Application.Worksheets("inicio").Range("H11").Value)
OZWcadEstaAberto = (Err.Number = 0)
Set zwApp = Nothing
Err.Clear
End Function

Sub ZWcadInstance()
Dim zAp As ZcadApplication

Dim ZWcadAberto As Boolean

ZWcadAberto = OZWcadEstaAberto()
If ZWcadAberto Then
Set zAp = GetObject(, Application.Worksheets("inicio").Range("H11").Value)
Else
Set zAp = CreateObject("ZWCAD.Application")
'ZWcadApp = GetObject(, Application.Worksheets("inicio").Range("H11").Value)

End If
zAp.Visible = True
'If Not ZWcadAberto Then zAp.Quit
Set zAp = Nothing
End Sub



Sub Example_Select()
    ' This example creates a selection set to select only line in the current drawing
    
    ' Create the selection set
    Dim objssline As ZWcadSelectionSet
    Set objssline = ThisDocument.SelectionSets.Add("TEST")
    
    ' Set the filter
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    
    FilterType(0) = 0
    FilterData(0) = "Line"
         
    ' Select all lines in the current drawing by filter
    objssline.Select zcSelectionSetAll, , , FilterType, FilterData
    
End Sub

