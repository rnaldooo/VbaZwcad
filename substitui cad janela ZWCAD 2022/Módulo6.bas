Attribute VB_Name = "Módulo6"
Dim iRow
 
Sub ListFolders()
 iRow = 15
 Sheets("inicio").Select
 Range("a1").Select
 Call ListMyFiles(Range("e3").Value, Range("e4").Value)
 End Sub
 
 'Look at the new file in the VBE menu Tools -> References... whether Microsoft Scripting Runtime is checked.
Sub ListMyFolders(mySourcePath As String, IncludeSubfolders As Boolean)
 Set MyObject = New Scripting.FileSystemObject
 Set MySource = MyObject.GetFolder(mySourcePath)
 On Error Resume Next
 For Each myFolder In MySource.Files 'SubFolder
 iCol = 2
 Cells(iRow, iCol).Value = myFolder.path
 iCol = iCol + 1
 Cells(iRow, iCol).Value = myFolder.Name
 iCol = iCol + 1
 Cells(iRow, iCol).Value = myFolder.DateLastModified
 iRow = iRow + 1
 Next
 If IncludeSubfolders Then
 For Each MySubFolder In MySource.SubFolders
 Call ListMyFolders(MySubFolder.path, True)
 Next
 End If
 End Sub


Sub ListMyFiles(mySourcePath As String, IncludeSubfolders As Boolean)
    Dim MyObject As Object
    Set MyObject = New Scripting.FileSystemObject
     Dim MySource As Folder
    Set MySource = MyObject.GetFolder(mySourcePath)
    On Error Resume Next
   Dim MyFile As File
   Dim iCol As Integer
     For Each MyFile In MySource.Files
       If InStr(MyFile.path, Sheets("inicio").Range("E5").Value) <> 0 Then
            iCol = 5
            Cells(iRow, iCol).Value = MyFile.path
            iCol = iCol + 1
            Cells(iRow, iCol).Value = MyFile.Name
            iCol = iCol + 1
            Cells(iRow, iCol).Value = MyFile.Size
            iCol = iCol + 1
            Cells(iRow, iCol).Value = MyFile.DateLastModified
            iRow = iRow + 1
            End If
         Next
    
    Columns("C:E").AutoFit
    Dim MySubFolder As Folder
    If IncludeSubfolders Then
        For Each MySubFolder In MySource.SubFolders
            Call ListMyFiles(MySubFolder.path, True)
        Next
    End If
End Sub

Sub Subst()
Dim ia, ib, ifa, ifb As Integer
Dim s1, s2 As String
ifa = Sheets("inicio").Range("e14").Value
ifb = Sheets("lista").Range("b2").Value
For ia = 1 To ifa
            For ib = 1 To ifb
            Call ReplaceTextInFile(Sheets("inicio").Range("E" & ia + 14 & "").Value, Sheets("lista").Range("a" & ib + 4 & "").Value, Sheets("lista").Range("b" & ib + 4 & "").Value)
            Next
Next
End Sub

