' Dim varray()

' For i = 1 to 10
'     Redim Preserve varray(i)
'     varray(i) = "valor" & i
' Next

' MsgBox(LBound(varray)) 'LBound Retorna o menor indece de um array'
' MsgBox(UBound(varray)) 'UBound Retorna o maior indece de um array'


' For i = LBound(varray) to UBound(varray)
'     MsgBox(varray(i))
' Next

' For Each el In varray
'     MsgBox (el)
' Next ' varray


' Não deu muito certo esse codigo das semanas'
' Mais tarde tento de outra forma'
' Para chegar no mesmo objetivo' 

Dim varray()


On Error Resume Next
contador = 1

Do While True
    Redim Preserve varray(contador)
    varray(cotador) = WeekdayName(contador)
    If Err.Number <> 0 Then
        Exit Do
    End If
    contador = contador + 1
Loop
On Error Goto 0

For Each el In varray
    MsgBox(el)
Next ' varray

'B = Filter(varray, "s", False) os que não tenha o s no array
'B = Filter(varray, "s", False, 0) os que não tenha o s no array 0 faz a comparação de texto e 1 faz a comparação binaria

B = Filter(varray, "-") 

For Each el In B
    MsgBox(el)
Next ' el

