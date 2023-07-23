'Operadores = , < menor, > maior, <> diferente'

a = InputBox("Informe o numero: ")

if CDbl(a) = 0 Then
    MsgBox("o numero e igual a zero")
ElseIf CDbl(a) > 0 Then
    MsgBox("o numero e maior que zero")
Else
    MsgBox("o numero e menor que zero")
End If