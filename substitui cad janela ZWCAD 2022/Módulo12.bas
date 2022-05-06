Attribute VB_Name = "Módulo12"
Function Avalia(ByVal s As String) As String

Avalia = Evaluate(s)

End Function


Function substlistaH(StrTexto As String, RangeValor As Range, RangeNovoValor As Range) As String

Dim rng1 As Variant, rng2 As Variant, str As String, itam As Integer

rng1 = RangeValor.Value
rng2 = RangeNovoValor.Value
str = StrTexto

itam = RangeValor.Count
    
    For i = 1 To itam
  '  MsgBox rng1(1, i)
    If rng1(1, i) = Empty Then
    rng1(1, i) = ""
    Else
    End If
    
    If rng2(1, i) = Empty Then
    rng2(1, i) = ""
    Else
    End If
    
    
    str = Replace(str, rng1(1, i), rng2(1, i)) ' linhas
    Next i
    
    substlistaH = str
End Function

Sub teste45()
Dim s, t As String
t = Range("M21").Value
s = substlistaV(t, Range("H15:H35"), Range("I15:I35"))
Range("f22").Value = s
End Sub
Function substlistaV(StrTexto As String, RangeValor As Range, RangeNovoValor As Range) As String

Dim rng1 As Variant, rng2 As Variant, str As String, itam As Integer

rng1 = RangeValor.Value
rng2 = RangeNovoValor.Value
str = StrTexto

itam = RangeValor.Count
    
    For i = 1 To itam
  '  MsgBox rng1(1, i)
    If rng1(i, 1) = Empty Then
    rng1(i, 1) = ""
    Else
    End If
    
    If rng2(i, 1) = Empty Then
    rng2(i, 1) = ""
    Else
    End If
    
    
    str = Replace(str, rng1(i, 1), rng2(i, 1)) ' linhas
    Next i
    
    substlistaV = str
End Function
