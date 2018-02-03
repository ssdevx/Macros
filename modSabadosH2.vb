Option Explicit
Type Checadas
    sabadoFecha As Date
End Type

Public excepcion(10) As Checadas

Public Sub iniciaExcepciones()
    Sheets("Hoja2").Select
    Dim j As Integer
    Dim i As Integer
    Dim strCol As String
    Const LIMITE = 10
    j = 3
    For i = 0 To LIMITE
        strCol = "K" + CStr(j + i)
        excepcion(i).sabadoFecha = Range(strCol).Value
    Next i
End Sub
