   Option Explicit

Sub limpiar()
    Columns("A:A").Select               ' Seleccionar toda la columna A
    Selection.Interior.ColorIndex = 0   ' Eliminar colores en las celdas
    Selection.Font.ColorIndex = 0       ' Cambiar color de la fuente
    
    Columns("B:B").Select
    Selection.Interior.ColorIndex = 0
    Selection.Font.ColorIndex = 0
    
    Columns("C:C").Select
    Selection.Interior.ColorIndex = 0
    Selection.Font.ColorIndex = 0
    
    Columns("D:D").Select
    Selection.Interior.ColorIndex = 0
    Selection.Font.ColorIndex = 0
        
    Columns("E:E").Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = 0
    Selection.Font.ColorIndex = 0
    
    Columns("F:F").Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = 0
    
    Columns("G:G").Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = 0
    Selection.Font.ColorIndex = 0
    
    '   Volver a transformar las cadenas, (MORELOS -> Morelos)
    
    Dim Row As Long
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Row = 1
    
    While ws.Cells(Row, 1).Value <> ""
        ws.Cells(Row, 4).Value = StrConv(ws.Cells(Row, 4).Value, vbProperCase)
         Row = Row + 1
    Wend

End Sub


