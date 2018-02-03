Option Explicit
Type Sabados
    sabadoFecha As Date
End Type

Public excepcion_sabado(10) As Sabados

Public Sub inicializaSabados()
    Sheets("Hoja3").Select
    Dim j As Integer
    Dim i As Integer
    Dim strCol As String
    Const LIMITE = 10
    j = 2
    For i = 0 To LIMITE
        strCol = "P" + CStr(j + i)
        excepcion_sabado(i).sabadoFecha = Range(strCol).Value
    Next i
End Sub
