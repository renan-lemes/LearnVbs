'testando o trim'
a = "    (Opa bao)   "

MsgBox (Trim(a))
MsgBox (ltrim(a))   'tirando os espaços a esquerda'
MsgBox (rtrim(a))   'tirando os espaços a direita '

MsgBox (ucase(Trim(a)))
a = UCase(Trim(a))
MsgBox (replace(a, "A", "b"))