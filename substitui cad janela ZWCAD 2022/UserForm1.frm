VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "modificar cad"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13620
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
ComboBox1.AddItem ("listar")
ComboBox1.AddItem ("deslocar")
ComboBox1.AddItem ("substituir")
ComboBox1.AddItem ("multiplicar")
ComboBox1.ListIndex = 0
ComboBox2.AddItem ("Janela")
ComboBox2.AddItem ("Tudo")
ComboBox2.ListIndex = 0


'PlantListBox.LinkedCell = "F10"  'If you want to link 1 cell. This will be the BoundColumn.
'Range("A20") = PlantListBox1.Column(0, PlantListBox.ListIndex) 'Column 1 value
'Range("A21") = PlantListBox1.Column(1, PlantListBox.ListIndex) 'Column 2 value

'PlantListBox.ListFillRange = "$C6:$D10"

 'PlantListBox.ListFillRange = "C6:D10"

 'PlantListBox.ListFillRange = "Sheet1!C6:D10"
 

 ' There will be five columns in the list box
  ListBox1.ColumnCount = 4

  ' The list box will be populated by range "A1:E4"
  ListBox1.RowSource = "inicio!e15:h30"
ListBox1.ListIndex = 0
  ' The value selected from the list box will go into cell A6
  'ListBox1.ControlSource = "e15"

  'Place the ListIndex into cell a6
  'ListBox1.BoundColumn = 0


End Sub
