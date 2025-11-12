Option Explicit

Public Sub LlenarHojaRegresionSimple(ws As Worksheet, resultado As RegressionResult, datosX As Variant, datosY As Variant)
    Dim i As Long, fila As Long
    Dim n As Long
    Dim yPred As Double, residuo As Double
    
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
    
    n = UBound(datosX, 1)
    
    With ws
        ' === ENCABEZADO PRINCIPAL ===
        .Cells(1, 1).Value = "ANÁLISIS DE REGRESIÓN LINEAL SIMPLE - AUDITORÍA"
        FormatearEncabezadoPrincipal.Range ("A1:F1")
        
        ' === INFORMACIÓN DEL MODELO ===
        .Cells(2, 1).Value = "INFORMACIÓN DEL MODELO"
        FormatearEncabezadoSeccion .Range("A2:F2")
        
        fila = 3
        .Cells(fila, 1).Value = "Fecha de análisis:"
        .Cells(fila, 2).Value = FormatoDeHoraUniversal(Now)
        .Cells(fila, 5).Value = "Número de observaciones:"
        .Cells(fila, 6).Value = n
        .Cells(fila, 6).NumberFormat = "#,##0"
        
        fila = fila + 1
        .Cells(fila, 1).Value = "Método:"
        .Cells(fila, 2).Value = "Mínimos Cuadrados Ordinarios (MCO)"
        .Cells(fila, 5).Value = "Grados de libertad:"
        .Cells(fila, 6).Value = resultado.DF_Residual
        
        ' === COEFICIENTES DEL MODELO ===
        fila = fila + 2
        .Cells(fila, 1).Value = "COEFICIENTES DE REGRESIÓN"
        FormatearEncabezadoSeccion .Range("A" & fila & ":F" & fila)
        
        fila = fila + 1
        .Cells(fila, 1).Value = "Parámetro"
        .Cells(fila, 2).Value = "Coeficiente"
        .Cells(fila, 3).Value = "Error Estándar"
        .Cells(fila, 4).Value = "Estadístico t"
        .Cells(fila, 5).Value = "Valor p"
        .Cells(fila, 6).Value = "Significancia"
        FormatearEncabezadoTabla .Range("A" & fila & ":F" & fila)
        
        fila = fila + 1
        .Cells(fila, 1).Value = "Intercepto (ß0)"
        .Cells(fila, 2).Value = resultado.Coefficients(1, 1)
        .Cells(fila, 2).NumberFormat = "0.0000"
        .Cells(fila, 3).Value = resultado.StandardErrors(1, 1)
        .Cells(fila, 3).NumberFormat = "0.0000"
        .Cells(fila, 4).Value = resultado.TStats(1, 1)
        .Cells(fila, 4).NumberFormat = "0.0000"
        .Cells(fila, 5).Value = resultado.PValues(1, 1)
        .Cells(fila, 5).NumberFormat = "0.0000"
        .Cells(fila, 6).Value = IIf(resultado.PValues(1, 1) < 0.05, "Significativo", "No significativo")
        
        fila = fila + 1
        .Cells(fila, 1).Value = "Pendiente (ß1)"
        .Cells(fila, 2).Value = resultado.Coefficients(2, 1)
        .Cells(fila, 2).NumberFormat = "0.0000"
        .Cells(fila, 3).Value = resultado.StandardErrors(2, 1)
        .Cells(fila, 3).NumberFormat = "0.0000"
        .Cells(fila, 4).Value = resultado.TStats(2, 1)
        .Cells(fila, 4).NumberFormat = "0.0000"
        .Cells(fila, 5).Value = resultado.PValues(2, 1)
        .Cells(fila, 5).NumberFormat = "0.0000"
        .Cells(fila, 6).Value = IIf(resultado.PValues(2, 1) < 0.05, "Significativo", "No significativo")
        
        ' === ECUACIÓN DEL MODELO ===
        fila = fila + 2
        .Cells(fila, 1).Value = "ECUACIÓN DE REGRESIÓN"
        FormatearEncabezadoSeccion .Range("A" & fila & ":F" & fila)
        
        fila = fila + 1
        .Cells(fila, 1).Value = "y = " & Format(resultado.Coefficients(1, 1), "0.0000") & " + " & _
                              Format(resultado.Coefficients(2, 1), "0.0000") & " * x"
        .Range("A" & fila & ":F" & fila).Merge
        .Cells(fila, 1).Font.Italic = True
        .Cells(fila, 1).Font.Size = 11
        .Cells(fila, 1).HorizontalAlignment = xlCenter
        
        ' === ESTADÍSTICAS DEL MODELO ===
        fila = fila + 2
        .Cells(fila, 1).Value = "ESTADÍSTICAS DEL MODELO"
        FormatearEncabezadoSeccion .Range("A" & fila & ":F" & fila)
        
        fila = fila + 1
        .Cells(fila, 1).Value = "R² (Coeficiente de determinación)"
        .Cells(fila, 2).Value = resultado.R2
        .Cells(fila, 2).NumberFormat = "0.0000"
        
        .Cells(fila, 5).Value = "Estadístico F"
        .Cells(fila, 6).Value = resultado.FStat
        .Cells(fila, 6).NumberFormat = "0.0000"
        
        fila = fila + 1
        .Cells(fila, 1).Value = "R² Ajustado"
        .Cells(fila, 2).Value = resultado.R2Adjusted
        .Cells(fila, 2).NumberFormat = "0.0000"
        
        .Cells(fila, 5).Value = "Valor p (F)"
        .Cells(fila, 6).Value = Application.WorksheetFunction.F_Dist_RT(resultado.FStat, resultado.DF_Regression, resultado.DF_Residual)
        .Cells(fila, 6).NumberFormat = "0.0000"
        
        fila = fila + 1
        .Cells(fila, 1).Value = "Error estándar de estimación"
        .Cells(fila, 2).Value = Sqr(resultado.MSE)
        .Cells(fila, 2).NumberFormat = "0.0000"
        
        ' === ANÁLISIS DE VARIANZA ===
        fila = fila + 2
        .Cells(fila, 1).Value = "ANÁLISIS DE VARIANZA (ANOVA)"
        FormatearEncabezadoSeccion .Range("A" & fila & ":F" & fila)
        
        fila = fila + 1
        .Cells(fila, 1).Value = "Fuente"
        .Cells(fila, 2).Value = "Suma de Cuadrados"
        .Cells(fila, 3).Value = "Grados de Libertad"
        .Cells(fila, 4).Value = "Cuadrado Medio"
        .Cells(fila, 5).Value = "Estadístico F"
        .Cells(fila, 6).Value = "Valor p"
        FormatearEncabezadoTabla .Range("A" & fila & ":F" & fila)
        
        fila = fila + 1
        .Cells(fila, 1).Value = "Regresión"
        .Cells(fila, 2).Value = resultado.SSR
        .Cells(fila, 2).NumberFormat = "0.0000"
        .Cells(fila, 3).Value = resultado.DF_Regression
        .Cells(fila, 4).Value = resultado.SSR / resultado.DF_Regression
        .Cells(fila, 4).NumberFormat = "0.0000"
        .Cells(fila, 5).Value = resultado.FStat
        .Cells(fila, 5).NumberFormat = "0.0000"
        .Cells(fila, 6).Value = Application.WorksheetFunction.F_Dist_RT(resultado.FStat, resultado.DF_Regression, resultado.DF_Residual)
        .Cells(fila, 6).NumberFormat = "0.0000"
        
        fila = fila + 1
        .Cells(fila, 1).Value = "Residual"
        .Cells(fila, 2).Value = resultado.SSE
        .Cells(fila, 2).NumberFormat = "0.0000"
        .Cells(fila, 3).Value = resultado.DF_Residual
        .Cells(fila, 4).Value = resultado.MSE
        .Cells(fila, 4).NumberFormat = "0.0000"
        
        fila = fila + 1
        .Cells(fila, 1).Value = "Total"
        .Cells(fila, 2).Value = resultado.SST
        .Cells(fila, 2).NumberFormat = "0.0000"
        .Cells(fila, 3).Value = n - 1
        
        ' === ANÁLISIS DE VARIANZA ===
        fila = fila + 2
        .Cells(fila, 1).Value = "ANÁLISIS COMPLE"
        FormatearEncabezadoSeccion .Range("A" & fila & ":F" & fila)
        
        fila = fila + 1
        .Cells(fila, 1).Value = "Total"
        .Cells(fila, 2).Value = resultado.SST
        .Cells(fila, 2).NumberFormat = "0.0000"
        .Cells(fila, 3).Value = n - 1
        
        ' Ajustar columnas
        .Columns("A:F").AutoFit
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


