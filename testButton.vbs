' ' Retorno do botao da Msgbox'
' ' 1 Ok '
' ' 2 Cancelar '
' ' 3 Anular '
' ' 4 Repetir '
' ' 5 Ignore '
' ' 6 Sim'
' ' 7 Nao '

' a = MsgBox ("Teste o botao", 1)
' MsgBox(a)

' a = MsgBox ("Teste o botao", 2)
' MsgBox(a)

' a = MsgBox ("Teste o botao", 3)
' MsgBox(a)

' a = MsgBox ("Teste o botao", 4)
' MsgBox(a)

' a = MsgBox ("Teste o botao", 5)
' MsgBox(a)

' a = MsgBox ("Teste o botao", 6)
' MsgBox(a)

a = MsgBox("Escolha uma das opcao", 4)

If CDbl(a) = 6 Then 
    MsgBox ("Voce precionou o Sim")
Else
    MsgBox ("Voce precionou o Nao")

End If
