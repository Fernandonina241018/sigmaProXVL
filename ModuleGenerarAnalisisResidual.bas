Public Sub GenerarAnalisisResidual(ws As Worksheet, datosX As Variant, datosY As Variant, resultado As RegressionResult)
    Dim i As Long, fila As Long
    Dim n As Long
    Dim yPred As Double, residuo As Double
    Dim sumaResiduos As Double, sumaResiduosCuad As Double
    Dim mediaResiduos As Double, desvResiduos As Double
    Dim wsFunc As WorksheetFunction
    Dim calcMode As XlCalculation
    
    calcMode = Application.Calculation
    
    ' DESACTIVAR durante procesamiento
    With Application
        .ScreenUpdating = False        ' No actualizar pantalla (CRÍTICO)
        .Calculation = xlCalculationManual  ' Desactivar cálculos automáticos
        .EnableEvents = False          ' Desactivar eventos
        .DisplayStatusBar = False      ' Ocultar barra de estado
    End With
    
    On Error GoTo Cleanup
    
    Set wsFunc = Application.WorksheetFunction
    
    n = UBound(datosX, 1)
    
    With ws
        ' Encontrar la última fila con datos
        fila = .Cells(.Rows.count, 1).End(xlUp).Row + 2
        
        ' === ANÁLISIS DE RESIDUOS ===
        .Cells(fila, 1).Value = "ANÁLISIS DE RESIDUOS"
        FormatearEncabezadoSeccion .Range("A" & fila & ":F" & fila)
        
        fila = fila + 1
        .Cells(fila, 1).Value = "Observaciones"
        .Cells(fila, 2).Value = "Ítems"
        .Cells(fila, 3).Value = "X"
        .Cells(fila, 4).Value = "Y Real"
        .Cells(fila, 5).Value = "Y Predicho"
        .Cells(fila, 6).Value = "Residuo"
        FormatearEncabezadoTabla .Range("A" & fila & ":F" & fila)
        
        sumaResiduos = 0
        sumaResiduosCuad = 0
        
        For i = 1 To n
            fila = fila + 1
            yPred = resultado.Coefficients(1, 1) + resultado.Coefficients(2, 1) * datosX(i, 1)
            residuo = datosY(i, 1) - yPred
            sumaResiduos = sumaResiduos + residuo
            sumaResiduosCuad = sumaResiduosCuad + residuo ^ 2
            
            .Cells(fila, 2).Value = i
            .Cells(fila, 3).Value = datosX(i, 1)
            .Cells(fila, 3).NumberFormat = "0.00"
            .Cells(fila, 4).Value = datosY(i, 1)
            .Cells(fila, 4).NumberFormat = "0.00"
            .Cells(fila, 5).Value = yPred
            .Cells(fila, 5).NumberFormat = "0.00"
            .Cells(fila, 6).Value = residuo
            .Cells(fila, 6).NumberFormat = "0.00"
        Next i
        
        ' Estadísticos de residuos
        fila = fila + 2
        .Cells(fila, 1).Value = "ESTADÍSTICOS DE RESIDUOS"
        FormatearEncabezadoSeccion .Range("A" & fila & ":F" & fila)
        
        fila = fila + 1
        .Cells(fila, 1).Value = "Media de residuos:"
        .Cells(fila, 6).Value = sumaResiduos / n
        .Cells(fila, 6).NumberFormat = "0.0000"
        
        fila = fila + 1
        .Cells(fila, 1).Value = "Desviación estándar de residuos:"
        .Cells(fila, 6).Value = Sqr(sumaResiduosCuad / (n - 2))
        .Cells(fila, 6).NumberFormat = "0.0000"
        
        ' Prueba de normalidad de residuos (simplificada)
        fila = fila + 1
        .Cells(fila, 1).Value = "Prueba de normalidad (Shapiro-Wilk):"
        
        ' Nota: En una implementación real, se necesitaría una función para Shapiro-Wilk
        .Cells(fila, 6).Value = "LastSWResult"
        
        ' Apartado de firmas
        fila = fila + 2
        .Cells(fila, 1).Value = "ESPACIO DE FIRMAS"
        FormatearEncabezadoSeccion .Range("A" & fila & ":F" & fila)
        
        fila = fila + 2
        .Cells(fila, 1).Value = "Realizado Por/Firma:"
        .Cells(fila, 3).Value = "Fecha:"
        
        .Cells(fila + 2, 1).Value = "Verificado Por/Firma:"
        .Cells(fila + 2, 3).Value = "Fecha:"
        
        ' Ajustar columnas
        .Columns("A:F").AutoFit
        .Columns("A:F").HorizontalAlignment = xlCenter
        .Columns("A:F").VerticalAlignment = xlCenter
    End With
    
    ' Definición de variables
    'Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rngBorders As Range
    
    ' Establecer la hoja de trabajo activa
    Set ws = ActiveSheet
    
    ' --- Lógica Dinámica ---
    ' 1. Encontrar la última fila usada en la Columna A.
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    nextRow = lastRow + 1
    
    ' 2. Encontrar la última columna usada en la Fila con más datos (asumiendo que tiene los encabezados).
    lastCol = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    
    ' 3. Definir el rango dinámico: Desde A1 hasta (última fila, última columna)
    Set rngBorders = ws.Range(ws.Cells(1, 1), ws.Cells(nextRow, lastCol))
    ' -------------------------
    
    ' Configuración de la ventana (se mantiene si es el objetivo)
    ActiveWindow.DisplayGridlines = False
    
    With rngBorders
        
        ' 4. Eliminar bordes diagonales
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        
        ' 5. Aplicar borde inferior
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = 0
        End With
        
        ' 6. Aplicar borde derecho
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = 0
        End With
        
    End With
    
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
    
End Sub
