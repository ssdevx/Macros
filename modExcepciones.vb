Option Explicit

Type Excepciones
    IDEmpleado As Integer
End Type

Public excepcion(50) As Excepciones

Public Sub inicializaExcepciones()
    Sheets("Hoja3").Select  'Seleccionar hoja
    
    Dim j As Integer
    Dim i As Integer
    Dim strCol As String
    Const LIMITE = 50
    j = 3
    For i = 0 To LIMITE
        strCol = "K" + CStr(j + i)
        excepcion(i).IDEmpleado = Range(strCol).Value
    Next i
End Sub
