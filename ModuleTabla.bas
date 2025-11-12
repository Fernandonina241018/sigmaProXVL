Sub GenerarTablaAleatoria()
    Dim ws As Worksheet
    Dim fila As Long, col As Long
    Dim valorMin As Double, valorMax As Double
    Dim celda As Range
    
    ' Crear nueva hoja
    Set ws = ActiveWorkbook.Sheets.Add
    ws.Name = "DatosAleatorios"
    
    ' Definir rango de valores
    valorMin = 120.5
    valorMax = 121.5
    
    ' Agregar encabezados
    For col = 1 To 5
        ws.Cells(1, col).Value = "T.A.T" & 4001389 + col ' Genera 4001390 a 4001394
        ws.Cells(1, col).Font.Bold = True
        ws.Cells(1, col).HorizontalAlignment = xlCenter
    Next col
    
    ' Generar datos
    For fila = 2 To 1251
        For col = 1 To 5
            Set celda = ws.Cells(fila, col)
            celda.Value = Round(Rnd() * (valorMax - valorMin) + valorMin, 2)
        Next col
    Next fila
    
    ' Ajustar ancho de columnas
    ws.Columns("A:E").AutoFit
End Sub

Sub EliminarHojasDePersonalXLSB()
    Dim ws As Worksheet
    Application.DisplayAlerts = False ' Evita confirmaciones
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Delete
    Next ws
    
    Application.DisplayAlerts = True
End Sub

Public Sub AjustarColumnasHojaActiva()
    
    ' Desactiva la actualización de pantalla para una ejecución más rápida
    Application.ScreenUpdating = False
    
    ' Opción 1 (Recomendada): Ajusta solo el rango con datos. MÁS RÁPIDO.
    ActiveSheet.UsedRange.Columns.AutoFit
    
    ' Opción 2 (Más completo): Ajusta TODAS las columnas de la hoja (A:XFD, etc.). MÁS LENTO.
    ' ActiveSheet.Cells.EntireColumn.AutoFit
    
    ' Reactiva la actualización de pantalla
    Application.ScreenUpdating = True
    
End Sub
