## código de la macro de suma

Sub CalcularSuma()
' CalcularSuma Macro
' Macro que calcula la suma a partir de variables

Dim n1, n2, suma As Integer
n1 = Val(InputBox("Ingrese el primer numero"))
n2 = Val(InputBox("Ingrese el segundo numero"))

suma = n1 + n2

Sheets("suma").Cells(2, 2).Value = n1
Sheets("suma").Cells(3, 2).Value = n2
Sheets("suma").Cells(4, 2).Value = suma


MsgBox ("El resultado es: " & suma)

End Sub
Sub Limpiar()
'
' Limpiar Macro
' Limpia el contenido de las celdas
'


Sheets("suma").Cells(2, 2).Value = ""
Sheets("suma").Cells(3, 2).Value = Empty
Sheets("suma").Cells(4, 2).Value = Clear


End Sub
## insertar filas 

Sub MacroInsertarFila()
'
' MacroInsertarFila Macro
' Inserta fila
'

'
    Range("B2").Select
    Sheets("GuardarOperacion").Select
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown
    Sheets("Suma").Select
End Sub
## primera forma de guardar

Sub MacroGuardar1()
'
' MacroGuardar1 Macro
' Macro que guarda los valores
'

'
    Sheets("GuardarOperacion").Select
    Range("A2").Select
    Sheets("Suma").Select
    Range("B2:B4").Select
    Selection.Copy
    Sheets("GuardarOperacion").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Range("A1").Select
    Sheets("Suma").Select
    Range("B2").Select
End Sub
## unir macros

Sub GuardarSuma()
'
' GuardarSuma Macro
'

'
MacroInsertarFila
MacroGuardar1
Limpiar



End Sub
## unir macros

Sub GuardarSuma2()
'
' GuardarSuma2 Macro
'

'
MacroGuardar2
Limpiar

End Sub
## Referencias relativas

Sub MacroGuardar2()
'
' MacroGuardar2 Macro
'

'
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Copy
    Sheets("GuardarOperacion").Select
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Sheets("Suma").Select
    Range("B2").Select
    Application.CutCopyMode = False
    
    Range("B3").Select
End Sub
