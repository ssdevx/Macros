Option Explicit

Type HorarioSalida2
    IDEmpleado As Integer
    FechaE As Date
End Type

Public horarioEspecial2(10) As HorarioSalida2
Public Sub inicializaHE2()
    Sheets("Hoja2").Select
    
    Dim j As Integer   ' Contador para Columna
    Dim i As Integer   ' Contador para Fila
    Dim strCol As String ' Nombre de la columna Actual
    Const LIMIT = 10
    j = 3
    
    For i = 0 To LIMIT
        strCol = "M" + CStr(j + i)
        horarioEspecial2(i).IDEmpleado = Range(strCol).Value
        strCol = "N" + CStr(j + i)
        horarioEspecial2(i).FechaE = Range(strCol).Value
    Next i
End Sub
