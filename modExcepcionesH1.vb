Option Explicit

Type Checadas
    IDEmpleado As Integer
End Type

Public excepcion(50) As Checadas

Public Sub iniciaExcepciones()
    Sheets("Hoja1").Select  'Seleccionar hoja
    
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
