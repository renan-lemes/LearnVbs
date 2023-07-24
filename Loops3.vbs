Dim varray()

For i = 1 to 7 
    Redim Preserve varray(i)
    varray(i) = WeekdayName(i)
Next

' For Each el In varray
'     MsgBox(el)
' Next ' forEach para os dias da semana

MsgBox ("A variavel varray: " & IsArray(varray)) 'Verifica se é um array'

MsgBox(Join(varray, "-"))
MsgBox(Join(varray, ";")) 'Para gerar csv usamos o ;'


MsgBox("O resultado do join e :" & IsArray(Join(varray)))

'Join transforma um array em um valor caracter
' Parametros -> array'
' Delimeter -> dos valores do array o default dele é o espaço ou seja a string vazia " ".
'

texto = Join(varray, "###")

varray2 = split(texto, "###")

For Each Element In varray2
    MsgBox(Element)
Next ' Element

' Split transforma um texto em array 
' Value texto que quer ser transformado     
' Delimeter os valores a serem delimitados
' count defini a quantidade de substrings a serem retornadas   
' compare que tipo de comparação 0 binaria e 1 texto

