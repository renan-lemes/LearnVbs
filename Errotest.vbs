'Tratando o erro em uma specifica linha '
' On Error Resume Next
' MsgBox(1/0)

' If Err.Number <> 0 Then
'     MsgBox("Erro tratado")
' End If
' On Error Goto 0

' MsgBox(1 / 0)

'-----------------------------------------'
a = InputBox("Digite um numero qualquer: ")
b = InputBox("Digite um outro qualquer: ")

On Error Resume Next
    
    c = a / b
    if Err.Number <> 0 Then
        MsgBox("Erro tratado")
    Else
        MsgBox(c)
    End If 

On Error Goto 0


