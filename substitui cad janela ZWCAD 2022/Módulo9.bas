Attribute VB_Name = "Módulo9"
Sub pegaponto()
'--------------------------------------------------------------------------------------------------------------------------------
' reinaldo - 25/03/2013  - pega ponto
'--------------------------------------------------------------------------------------------------------------------------------
' Dim w As Window
'Set w = Application.ActiveWindow

'Dim i, ideci As Integer
'Dim vtemp As Variant

' AUTOCAD**** - definindo desenho aberto
Set ZWcad1 = GetObject(, "ZWCAD.Application.2020")
Dim aAD As ZWCAD.ZcadDocument
Set aAD = ZWCAD.ActiveDocument
aAD.Activate

Dim Excel As Object
Set Excel = GetObject(, "Excel.Application")                ' define excel como o objeto
Excel.Visible = False
'aAD.Active


AppActivate ZWcad1.Caption

 'Dim AutoCADAppID
 'Dim ZWcadApp As ZWcadApplication
  
 'Set ZWcadApp = GetObject(, "AutoCAD.Application")
 'AutoCADAppID = ZWcadApp.Caption
 'AppActivate AutoCADAppID

'ZWcad.Visible = True


On Error GoTo mostrar
Dim Vponto01, Vponto02 As Variant
Dim aADu As ZcadUtility
Set aADu = aAD.Utility
    aADu.Prompt (Chr(10))
    aADu.Prompt ("SELECIONE O PONTO 1)")
    Vponto01 = aADu.GetPoint(, "selecione:")

aADu.Prompt (Chr(10))
    aADu.Prompt ("SELECIONE O PONTO 1)")
    Vponto02 = aADu.GetPoint(, "selecione:")

Worksheets("inicio").Range("G7").Value = Vponto01(0)
Worksheets("inicio").Range("H7").Value = Vponto01(1)
Worksheets("inicio").Range("G8").Value = Vponto02(0)
Worksheets("inicio").Range("H8").Value = Vponto02(1)

Excel.Visible = True
AppActivate Application.Caption

aAD.Regen zcAllViewports
aAD.ActiveViewport.ZoomExtents

mostrar:
Excel.Visible = True


 'AppActivate Excel.Caption, 0
 'Dim tit As String
 'tit = Application.Caption
 

 'AppActivate2 w.Caption
' AppActivate w.Caption

 
 'AppActivate tit
 
 'Excel.getfocus

End Sub
