' a = "10"
' b = CInt(a)

' c = Asc(b)
' MsgBox (b)
' MsgBox (c)


' a = "10,123"

' c = CDbl(a) + 1
' MsgBox (c)
' MsgBox (CLng(a) + 3) 'tipo de dado que é inteiro '
' MsgBox (CSng(a) + 4)
' MsgBox (CCur(a)) ' tipo moeda '

'----------------------------------------------------------'

a = CDate("19/11/1996 10:00:00") + (1/24)
b = DateValue(a)
c = TimeValue(a)
MsgBox (a)
MsgBox (b)
MsgBox (c)