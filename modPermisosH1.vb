Option Explicit

Type HorarioSalida
    IDEmpleado As Integer
    FechaE As Date
End Type

Public horarioEspecial(50) As HorarioSalida
Public Sub inicializaHE()
    Sheets("Hoja1").Select  'Seleccionar la hoja
    
    Dim j As Integer   ' Contador para Columna
    Dim i As Integer   ' Contador para renglon
    Dim strCol As String    'Fila
    Const LIMIT_SUP = 50
    j = 3 'Los datos empiezan a capturar desde la fila 3
    
    For i = 0 To LIMIT_SUP
       strCol = "M" + CStr(j + i)
       horarioEspecial(i).IDEmpleado = Range(strCol).Value 'Capturar el ID del empleado
       strCol = "N" + CStr(j + i)
       horarioEspecial(i).FechaE = Range(strCol).Value  'Capturar la fecha de horario especial
    Next i
End Sub
