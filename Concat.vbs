' Concatenação de strings podemos usar tanto o + quanto o &'
' Consegue cortar a string através do Mid (valor, inicio, fim) '
a = InputBox ("Informe o o valor a")
b = InputBox ("informe o valor b")

MsgBox ("temos " + a & b)

c = Mid(a & b, 1, 3) 

MsgBox ("result Mid " + c)