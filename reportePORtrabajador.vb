Option Explicit

Sub reportePORtrabajador()
    Dim Row As Long             'Recorrer filas
    Dim ws As Worksheet         'Hoja activa
    Dim contador As Integer     'Contador (1, 2, 3 ó 4 checadas)
    Dim tolerancia As Integer   'Tolerancia para horas de comida
    Dim r1 As Integer           'Recorrer excepciones
    Dim r2 As Integer           'Recorrer sabados
    Dim fecha As Date           'Fecha (sabados)
    Dim minutos As Date         'Minutos (comida)
    Dim fecha1 As String        'Fecha 1
    Dim fecha2 As String        'Fecha 2
    Dim uChecada As String      'Ultima checada (00:00:00)
    Dim uChecada_1 As String    'Ultima checada (00:00)
    Dim hPermiso As String      'Hora permiso
    Dim conteo As Integer       'Controlador de tiempos
   
    Dim hTotal As Date          'Variables para hora total
    Dim hEntrada As Date        'Hora de entrada
    Dim hSalida As Date         'Hora de salida
    Dim hora As Double          'Almacenar hora
    Dim min As Double           'Almacenar minutos
    Dim seg As Double           'Almacenar segundos
    Dim cht As String           'Almacenar cadena ( 00 horas   00 minutos)
    
    Dim MAX_EXCEPCION As Integer
    Dim MAX_PERMISOS As Integer
    
    Worksheets("Hoja2").Activate    'Activar hoja2
    tolerancia = Range("I1").Value  'Almacenar valor tolerancia I1
    Set ws = ActiveSheet
    Row = 1
    Const MAX = 60
    MAX_EXCEPCION = 10
    MAX_PERMISOS = 10
    contador = 0
    conteo = 0
    
    modSabadosH2.iniciaExcepciones
    modPermisosH2.inicializaHE2
    
    While ws.Cells(Row, 3).Value <> ""
        'Convertir en mayusculas el nombre las sucursales, columna D
         ws.Cells(Row, 4).Value = UCase(ws.Cells(Row, 4).Value)
    
    
        fecha1 = Format(ws.Cells(Row, 3).Value, "Short Date")
        fecha2 = Format(ws.Cells(Row + 1, 3).Value, "Short Date")
        
        If fecha1 = fecha2 Then
            contador = contador + 1
        Else
            Select Case contador
            Case 0
                Range("A" & (Row), "E" & (Row)).Interior.Color = RGB(255, 96, 96)
            Case 1
                '--------------------------------------------------------------------------------------------------------
                'Calcular total de horas de trabajo.
                ' hEntrada   ->  Horario de entrada
                ' hSalida    ->  Horario de salida
                ' hTotal -> Diferencia entre Horario de entrada y horario de salida, en segundos.
                ' Aplicar algoritmo de conversion para sacar horas, min y seg apartir de los segundos(hTotal)
                hEntrada = Range("C" & (Row - 1)).Value
                hSalida = Range("C" & (Row)).Value
                hTotal = CDbl(DateDiff("s", hEntrada, hSalida))
                
                '_-_Algortimo
                hora = Fix(hTotal / 3600)
                min = Fix((hTotal - (hora * 3600)) / 60)
                seg = hTotal - (hora * 3600) - (min * 60)
                
                'Contar las horas "extras" del trabajador, a partir de las siguientes condiciones.
                '   10 horas y 45 minutos   ----->  1 hora
                '   11 horas y 45 minutos   ----->  2 horas
                '   12 horas y 45 minutos   ----->  3 horas
                '   ...
                '   16 horas y 45 minutos   ----->  7 horas
                If hTotal >= 38700 And hTotal <= 42240 Then
                    conteo = 1
                ElseIf hTotal >= 42300 And hTotal <= 45840 Then
                    conteo = 2
                ElseIf hTotal >= 45900 And hTotal <= 49440 Then
                    conteo = 3
                ElseIf hTotal >= 49500 And hTotal <= 53040 Then
                    conteo = 4
                ElseIf hTotal >= 53100 And hTotal <= 57240 Then
                    conteo = 5
                ElseIf hTotal >= 57300 And hTotal <= 60240 Then
                    conteo = 6
                ElseIf hTotal >= 60300 And hTotal <= 63840 Then
                    conteo = 7
                Else
                    conteo = 0
                End If
                
                'Armar cadena, imprimirlo en columna F y dar formato para que se visualice bien.
                cht = (Trim(hora) + " Horas   " + Trim(min) + " Minutos")
                Range("F" & (Row)).Value = cht
                Range("F" & (Row)).NumberFormat = "General"
                Range("G" & (Row)).Value = conteo
                Range("G" & (Row)).NumberFormat = "General"
                
                'Pintar renglones cuando horas activas sean mayor o igual a 10hrs y 45min.
                If hTotal >= 38700 Then
                    Range("F" & (Row)).Interior.ColorIndex = 8
                End If
                'Pintar las celdas donde conteo(horas extras) sea mayor a 0
                If conteo > 0 Then
                    Range("G" & (Row)).Interior.ColorIndex = 8
                End If
                '-----------------------------------------------------------------------------------------------------------
                
                '===========================================================================================================
                ' Recorre la base de datos para ver si la fecha de checada es sabado
                ' Si la fecha es sabado coloca la palabra SABADO en las dos filas del registro
                For r1 = 0 To MAX_EXCEPCION
                    fecha = Format(ws.Cells(Row, 3).Value, "Short Date")
                    If (modSabadosH2.excepcion(r1).sabadoFecha <> fecha) Then
                        Range("A" & (Row - 1), "E" & (Row - 1)).Interior.Color = RGB(255, 204, 0)
                        Range("A" & (Row), "E" & (Row)).Interior.Color = RGB(255, 204, 0)
                    Else
                        Range("A" & (Row - 1), "E" & (Row - 1)).Interior.ColorIndex = 0
                        Range("A" & (Row), "E" & (Row)).Interior.ColorIndex = 0
                        Range("E" & (Row - 1)).Value = "SABADO"
                        Range("E" & (Row - 1)).Font.ColorIndex = 3
                        Range("E" & (Row - 1)).HorizontalAlignment = xlRight
                        Range("E" & (Row)).Value = "SABADO"
                        Range("E" & (Row)).Font.ColorIndex = 3
                        Range("E" & (Row)).HorizontalAlignment = xlRight
                        r1 = MAX_EXCEPCION
                    End If
                Next r1
                '===========================================================================================================
                
            Case 2
                'Solo Pintar los reglones para la verificación manual del usuario.
                Range("A" & (Row - 2), "E" & (Row - 2)).Interior.Color = RGB(255, 204, 0)
                Range("A" & (Row - 1), "E" & (Row - 1)).Interior.Color = RGB(255, 204, 0)
                Range("A" & (Row), "E" & (Row)).Interior.Color = RGB(255, 204, 0)
            Case 3
                '------------------------------------------------------------------------------------------------------------
                ' Calcular diferencia entre horas de comida
                minutos = DateDiff("n", ws.Cells(Row - 2, 3).Value, ws.Cells(Row - 1, 3).Value)
                Range("E" & (Row - 1)).Value = minutos
                Range("E" & (Row - 1)).NumberFormat = "General"
                If minutos > (MAX + tolerancia) Then        'Si minutos rebasa a 60 + tolerancia(consideracion del usuario), pinta como retardo.
                    Range("A" & (Row - 2), "E" & (Row - 2)).Interior.Color = RGB(255, 153, 0)
                    Range("A" & (Row - 1), "E" & (Row - 1)).Interior.Color = RGB(255, 153, 0)
                End If
                '------------------------------------------------------------------------------------------------------------
                
                '============================================================================================================
                'Calcular total de horas de trabajo.
                hEntrada = Range("C" & (Row - 3)).Value         ' hEntrada   ->  Horario de entrada
                hSalida = Range("C" & (Row)).Value              ' hSalida    ->  Horario de salida
                hTotal = CDbl(DateDiff("s", hEntrada, hSalida)) ' hTotal -> Diferencia entre Horario de entrada y horario de salida, en segundos.
                ' Aplicar algoritmo de conversion para sacar horas, min y seg apartir de los segundos(hTotal)
                hora = Fix(hTotal / 3600)
                min = Fix((hTotal - (hora * 3600)) / 60)
                seg = hTotal - (hora * 3600) - (min * 60)
                
                'Contar las horas "extras" del trabajador, a partir de las siguientes condiciones.
                '   10 horas y 45 minutos   ----->  1 hora
                '   11 horas y 45 minutos   ----->  2 horas
                '   12 horas y 45 minutos   ----->  3 horas
                '   ...
                '   16 horas y 45 minutos   ----->  7 horas
                If hTotal >= 38700 And hTotal <= 42240 Then
                    conteo = 1
                ElseIf hTotal >= 42300 And hTotal <= 45840 Then
                    conteo = 2
                ElseIf hTotal >= 45900 And hTotal <= 49440 Then
                    conteo = 3
                ElseIf hTotal >= 49500 And hTotal <= 53040 Then
                    conteo = 4
                ElseIf hTotal >= 53100 And hTotal <= 57240 Then
                    conteo = 5
                ElseIf hTotal >= 57300 And hTotal <= 60240 Then
                    conteo = 6
                ElseIf hTotal >= 60300 And hTotal <= 63840 Then
                    conteo = 7
                Else
                    conteo = 0
                End If
                
                cht = (Trim(hora) + " Horas   " + Trim(min) + " Minutos")   'Cadena final con Horas y Minutos calculados
                Range("F" & (Row)).Value = cht                              'Imprimir cth(cadena hora total) en Columna F
                Range("F" & (Row)).NumberFormat = "General"
                Range("G" & (Row)).Value = conteo
                Range("G" & (Row)).NumberFormat = "General"
                
                'Pintar renglones cuando horas activas sean mayor o igual a 10hrs y 45min.
                If hTotal >= 38700 Then
                    Range("F" & (Row)).Interior.ColorIndex = 8
                End If
                'Pintar las celdas donde conteo(horas extras) sea mayor a 0
                If conteo > 0 Then
                    Range("G" & (Row)).Interior.ColorIndex = 8
                End If
                '============================================================================================================
                
                
                '------------------------------------------------------------------------------------------------------------
                'Verificar horas de salida.
                uChecada = CDate(Format(ws.Cells(Row, 3).Value, "hh:mm:ss"))    'Extraer ultima hora checada en formato 00:00:00
                uChecada_1 = CDate(Format(ws.Cells(Row, 3).Value, "hh:mm"))     'Extraer ultima hora checada en formato 00:00
                'Comparar si la última checada es antes de las 19:00 horas
                If uChecada < CDate("19:00:00") Then
                    'Recorrer si tiene permisos de hacer ultima checada antes de las 19:00hras
                    'Si hay permiso de checar(salir) antes de la 19:00, colocar la palabra EXCEPCION 2
                    For r2 = 0 To MAX_PERMISOS
                        hPermiso = Format(modPermisosH2.horarioEspecial2(r2).FechaE, "hh:mm")
                        If (modPermisosH2.horarioEspecial2(r2).IDEmpleado = ws.Cells(Row, 1).Value And uChecada_1 >= hPermiso) Then
                            Range("A" & (Row), "E" & (Row)).Interior.ColorIndex = 0
                            Range("E" & (Row)).Value = "EXCEPCION 2"
                            Range("E" & (Row)).Font.ColorIndex = 33
                            Range("E" & (Row)).HorizontalAlignment = xlRight
                            r2 = MAX_PERMISOS
                        Else    'Sino tiene permisos, pintar infracción
                            Range("A" & (Row), "E" & (Row)).Interior.Color = RGB(135, 206, 235)
                        End If
                    Next r2
                End If
                '------------------------------------------------------------------------------------------------------------
            End Select
            contador = 0
        End If
        Row = Row + 1
    Wend
End Sub
