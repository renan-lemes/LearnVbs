a = 0.5

MsgBox (FormatNumber(a))
MsgBox(FormatNumber(a, 3,-1))
MsgBox(FormatNumber(a, 3,0))

a = 20530.53

MsgBox (FormatNumber(a, 2, , ,-1))
MsgBox (FormatNumber(a, 2, , ,0))

a = -1020530.53

MsgBox (FormatNumber(a, 2, ,-1,-1))
MsgBox (FormatNumber(a, 2, ,0,0))

'-------------------------------------'
a = 0.5

MsgBox (FormatCurrency(a))
MsgBox(FormatCurrency(a, 3,-1))
MsgBox(FormatCurrency(a, 3,0))

a = 20530.53

MsgBox (FormatCurrency(a, 2, , ,-1))
MsgBox (FormatCurrency(a, 2, , ,0))

a = -1020530.53

MsgBox (FormatCurrency(a, 2, ,-1,-1))
MsgBox (FormatCurrency(a, 2, ,0,0))

'-------------------------------------'

a = 0.5

MsgBox (FormatPercent(a))
MsgBox(FormatPercent(a, 3,-1))
MsgBox(FormatPercent(a, 3,0))

a = 20530.53

MsgBox (FormatPercent(a, 2, , ,-1))
MsgBox (FormatPercent(a, 2, , ,0))

a = -1020530.53

MsgBox (FormatPercent(a, 2, ,-1,-1))
MsgBox (FormatPercent(a, 2, ,0,0))