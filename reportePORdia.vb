Option Explicit
Sub reportePORdia()
    Dim clave1 As Integer       'Almacena clave 1
    Dim clave2 As Integer       'Almacena clave 2
    Dim contador As Integer     'Contador (1, 2, 3 o 4 checadas)
    Dim r1 As Integer           'Rango para recorrer excepciones
    Dim r2 As Integer           'Rango para recorrer horarios especiales
    Dim hTotal As Date          'Almacena Horas totales
    Dim tolerancia As Integer   'Almacena minutos tolerancia (comida)
    Dim uChecada As String      'Almacena última checada (00:00:00)
    Dim uChecada_1 As String    'Almacena última checada (00:00)
    Dim hora1 As Date           'Almacena hora salida a comer
    Dim hora2 As Date           'Almacena hora regreso a comer
    Dim minComida As Date       'Almacena minutos de hora comida
    Dim hPermiso As Date        'Almacena horas de permisos
    Dim MAX_EXCEPCIONES As Integer  'Numero máximo de excepciones
    Dim MAX_PERMISOS As Integer     'Numero máximo de permisos
    Dim hEntrada As Date    'Horario de entrada
    Dim hSalida As Date     'Horario de salida
    Dim hora As Double      'Almacena horas
    Dim min As Double       'Almacena minutos
    Dim seg As Double       'Almacena segundos
    Dim cht As String       'Almacena cadena (xx horas   xx minutos)
    Dim conteo As Integer   'Controlador de tiempos ("horas extras")
        
    Dim Row As Long
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Row = 1
    contador = 0
    MAX_EXCEPCIONES = 50
    MAX_PERMISOS = 50
    Const MAX = 60
    tolerancia = Range("I1").Value
    
    modExcepcionesH1.iniciaExcepciones
    modPermisosH1.inicializaHE
    
    While ws.Cells(Row, 1).Value <> ""
        'Convertir en mayusculas el nombre de las sucursales, columna D
        ws.Cells(Row, 4).Value = UCase(ws.Cells(Row, 4).Value)
        
        
        clave1 = ws.Cells(Row, 1).Value
        clave2 = ws.Cells(Row + 1, 1).Value
        
        If clave1 = clave2 Then
            contador = contador + 1
        Else
            Select Case contador
            Case 0 '    Solo 1 checada
                Range("A" & (Row), "E" & (Row)).Interior.Color = RGB(255, 96, 96)
            
            Case 1 '    2 Checadas
                '--------------------------------------------------------------------------------------------------------
                'Calcular total de horas de trabajo.
                ' hEntrada   ->  Horario de entrada
                ' hSalida    ->  Horario de salida
                ' hTotal -> Diferencia entre Horario de entrada y horario de salida, en segundos.
                ' Aplicar algoritmo de conversion para sacar horas, min y seg apartir de los segundos(hTotal)
                hEntrada = Range("C" & (Row - 1)).Value
                hSalida = Range("C" & (Row)).Value
                hTotal = CDbl(DateDiff("s", hEntrada, hSalida))
                
                '_-_Algoritmo
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
                Range("G" & (Row - 1)).NumberFormat = "General"
                
                
                'Pintar renglones cuando horas activas sean mayor o igual a 10hrs y 45min.
                If hTotal >= 38700 Then
                    Range("F" & (Row)).Interior.ColorIndex = 8
                End If
                'Pintar si conteo ("Horas extras") sea mayor que 0
                If conteo > 0 Then
                    Range("G" & (Row)).Interior.ColorIndex = 8
                End If
                
                '===========================================================================================================
                'Exepciones para 2 checadas
                '-_- Recorre la base de datos de excepciones y para comparar si el trabajador tiene permiso de hacer solo 2 checada
                '-_- Si tiene permisos colocar EXCEPCION 1, sino, pintar las filas
                For r1 = 0 To MAX_EXCEPCIONES
                    If (modExcepcionesH1.excepcion(r1).IDEmpleado <> ws.Cells(Row, 1).Value) Then
                        Range("A" & (Row - 1), "E" & (Row - 1)).Interior.Color = RGB(255, 204, 0)
                        Range("A" & (Row), "E" & (Row)).Interior.Color = RGB(255, 204, 0)
                    Else
                        Range("A" & (Row - 1), "E" & (Row - 1)).Interior.ColorIndex = 0
                        Range("A" & (Row), "E" & (Row)).Interior.ColorIndex = 0
                        Range("E" & (Row)).Value = "EXCEPCION 1"
                        Range("E" & (Row)).Font.ColorIndex = 3
                        Range("E" & (Row)).HorizontalAlignment = xlRight
                        r1 = MAX_EXCEPCIONES
                    End If
                Next r1
                '===========================================================================================================
            
            Case 2 '    3 Checadas
                Range("A" & (Row - 2), "E" & (Row - 2)).Interior.Color = RGB(255, 204, 0)
                Range("A" & (Row - 1), "E" & (Row - 1)).Interior.Color = RGB(255, 204, 0)
                Range("A" & (Row), "E" & (Row)).Interior.Color = RGB(255, 204, 0)
            
            Case 3  '   4 Checadas
                '------------------------------------------------------------------------------------------------------------
                ' Calcular diferencia entre horas de comida
                hora1 = Range("C" & (Row - 2)).Value    ' hora1 ->  Horario Salida(comida)
                hora2 = Range("C" & (Row - 1)).Value    ' hora2 ->  Horario regreso(comida)
                minComida = DateDiff("n", hora1, hora2) ' minComida ->  Diferecia entre hora1 y hora2, refeljado en minutos.
                Range("E" & (Row - 1)).Value = minComida    ' Imprimir valor de minComida en Columna E
                Range("E" & (Row - 1)).NumberFormat = "General"
                
                If minComida > (MAX + tolerancia) Then  'Si minComida rebasa a 60 + tolerancia(consideracion del usuario), pinta como retardo.
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
                
                cht = (Trim(hora) + " Horas   " + Trim(min) + " Minutos") 'Cadena final con Horas y Minutos calculados
                Range("F" & (Row)).Value = cht                              'Imprimir cth(cadena hora total) en Columna F
                Range("F" & (Row)).NumberFormat = "General"
                Range("G" & (Row)).Value = conteo
                Range("G" & (Row - 1)).NumberFormat = "General"
                
                'Pintar renglones cuando horas activas sean mayor o igual a 10hrs y 45min.
                If hTotal >= 38700 Then
                    Range("F" & (Row)).Interior.ColorIndex = 8
                End If
                
                If conteo > 0 Then
                    Range("G" & (Row)).Interior.ColorIndex = 8
                End If
                '============================================================================================================
                  
                uChecada = CDate(Format(ws.Cells(Row, 3).Value, "hh:mm:ss"))    'Extraer ultima hora checada en formato 00:00:00
                uChecada_1 = CDate(Format(ws.Cells(Row, 3).Value, "hh:mm"))     'Extraer ultima hora checada en formato 00:00
                'Comparar si la última checada es antes de las 19:00 horas
                If uChecada < CDate("19:00:00") Then
                'Recorrer si tiene permisos de hacer ultima checada antes de las 19:00hras
                'Si hay permiso de checar(salir) antes de la 19:00, colocar la palabra EXCEPCION 2
                    For r2 = 0 To MAX_PERMISOS
                        hPermiso = Format(modPermisosH1.horarioEspecial(r2).FechaE, "hh:mm")
                        If (modPermisosH1.horarioEspecial(r2).IDEmpleado = ws.Cells(Row, 1).Value) And uChecada_1 >= hPermiso Then
                            Range("A" & (Row), "E" & (Row)).Interior.ColorIndex = 0
                            Range("E" & (Row)).Value = "EXCEPCION 2"
                            Range("E" & (Row)).Font.ColorIndex = 33
                            Range("E" & (Row)).HorizontalAlignment = xlRight
                            r2 = MAX_PERMISOS
                        Else    'Sino tiene permisos, pintar la infracción para verificacion manual del usuario.
                            Range("A" & (Row), "E" & (Row)).Interior.Color = RGB(135, 206, 235)
                        End If
                    Next r2
                End If
            End Select
            contador = 0
        End If
        Row = Row + 1
    Wend
End Sub

