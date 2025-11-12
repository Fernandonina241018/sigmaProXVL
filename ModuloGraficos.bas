Option Explicit

' Constantes para configuración de gráficos - AUMENTADAS para mejor lectura
Private Const ANCHO_GRAFICO As Double = 450    ' ?? Aumentado de 350
Private Const ALTO_GRAFICO As Double = 300     ' ?? Aumentado de 200
Private Const ESPACIO_ENTRE_GRAFICOS As Long = -18  ' ?? REDUCIDO de 15 para acercar gráficos

' Función principal para crear todos los gráficos de una columna
Public Sub CrearGraficosParaColumna(ws As Worksheet, stats As EstadisticasColumna, topPosition As Long, leftColumn As Integer)
    Dim chartTop As Long
    chartTop = topPosition
    
    ' 1. Crear gráfico de control
    chartTop = CrearGraficoControl(ws, stats, chartTop, leftColumn)
    
    ' 2. Crear gráfico de dispersión debajo del de control
    ' ?? ESPACIO REDUCIDO: Solo 5 puntos entre gráficos (antes 15)
    chartTop = chartTop + (ALTO_GRAFICO / 15) + ESPACIO_ENTRE_GRAFICOS
    chartTop = CrearGraficoDispersion(ws, stats, chartTop, leftColumn)
    
    ' 3. Crear histograma con campana de Gauss
    chartTop = chartTop + (ALTO_GRAFICO / 15) + ESPACIO_ENTRE_GRAFICOS
    Call CrearHistogramaConGauss(ws, stats, chartTop, leftColumn)
End Sub

' Función para crear gráfico de control con línea de tendencia
Private Function CrearGraficoControl(ws As Worksheet, stats As EstadisticasColumna, topPosition As Long, leftColumn As Integer) As Long
    Dim cht As ChartObject
    Dim serie As Series
    Dim tendencia As Trendline
    Dim i As Long
    Dim xValores() As Variant
    Dim yValores() As Variant
    
    Dim limiteSuperior As Double
    Dim limiteInferior As Double
    
    limiteSuperior = CDbl(sigmaproxvl.cboLimiteSuperior.Value)
    limiteInferior = CDbl(sigmaproxvl.cboLimiteInferior.Value)

    Dim yLimSup() As Double, yLimInf() As Double
    ReDim yLimSup(1 To stats.count)
    ReDim yLimInf(1 To stats.count)
    
    For i = 1 To stats.count
        yLimSup(i) = limiteSuperior
        yLimInf(i) = limiteInferior
    Next i

    ' Preparar datos para el gráfico de control (secuencia temporal)
    ReDim xValores(1 To stats.count)
    ReDim yValores(1 To stats.count)
    
    For i = 1 To stats.count
        xValores(i) = i ' Eje X: número de medición
        yValores(i) = stats.valores(i - 1) ' Eje Y: valor de la medición
    Next i
    
    ' ?? GRÁFICO MÁS GRANDE: Width=400, Height=250 (antes 350x200)
    Set cht = ws.ChartObjects.Add(Left:=ws.Cells(topPosition, leftColumn).Left, _
                                  Top:=ws.Cells(topPosition, leftColumn).Top, _
                                  Width:=ANCHO_GRAFICO, _
                                  Height:=ALTO_GRAFICO)
    
    With cht.Chart
        .ChartType = xlXYScatterLines
        .HasTitle = True
        .chartTitle.Text = "Gráfico de Control - " & stats.NombreColumna
        .chartTitle.Font.Size = 11  ' ?? Ligeramente más grande para mejor lectura
        .chartTitle.Font.Bold = True
        
        ' Configurar ejes
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Número de Medición"
            .AxisTitle.Font.Size = 9  ' ?? Más grande
        End With
        
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Valor"
            .AxisTitle.Font.Size = 9  ' ?? Más grande
        End With
        
        ' Añadir serie de datos
        Set serie = .SeriesCollection.NewSeries
        With serie
            .Name = "Datos de " & stats.NombreColumna
            .XValues = xValores
            .Values = yValores
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 5  ' ?? Marcadores ligeramente más grandes
            .Format.Line.Weight = 1.5
        End With
        
        ' Añadir línea de tendencia lineal con ecuación
        If stats.count >= 3 Then
            Set tendencia = serie.Trendlines.Add
            With tendencia
                .Type = xlLinear
                .DisplayEquation = True
                .DisplayRSquared = True
                
                With .DataLabel
                    .Font.Size = 9
                    
                    ' Reposicionar manualmente
                    .Left = .Left + 22   ' Mueve hacia la derecha
                    .Top = .Top - 50     ' Mueve hacia abajo
                End With
            End With
        End If
        
        ' Línea de límite superior
        With .SeriesCollection.NewSeries
            .Name = "Límite Superior"
            .XValues = xValores
            .Values = yLimSup
            .Format.Line.ForeColor.RGB = RGB(255, 0, 0) ' Rojo
            .Format.Line.DashStyle = msoLineDash
            .Format.Line.Weight = 1
        End With
        
        ' Línea de límite inferior
        With .SeriesCollection.NewSeries
            .Name = "Límite Inferior"
            .XValues = xValores
            .Values = yLimInf
            .Format.Line.ForeColor.RGB = RGB(0, 0, 255) ' Azul
            .Format.Line.DashStyle = msoLineDash
            .Format.Line.Weight = 1
        End With

        ' Añadir líneas de control (media ± desviación estándar)
        If stats.count > 1 Then
            AddLineasControl cht.Chart, stats.promedio, stats.desviacionEstandar, stats.count, stats.mediana, stats.varianza, stats.moda, stats.asimetria
        End If
        
        ' Formato del gráfico
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(240, 240, 240)
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 9  ' ?? Leyenda más legible
    End With
    
    ' Devolver la posición inferior del gráfico
    CrearGraficoControl = topPosition + (ALTO_GRAFICO / 15)
End Function

' Función para añadir líneas de control al gráfico
Private Sub AddLineasControl(cht As Chart, promedio As Double, desviacion As Double, count As Long, mediana As Double, varianza As Double, moda As Double, asimetria As Double)
    Dim serie As Series
    Dim xValores() As Variant
    Dim yValores() As Variant
    
    ' Preparar datos para las líneas de control
    ReDim xValores(1 To 2)
    ReDim yValores(1 To 2)
    
    ' Línea de promedio
    xValores(1) = 1
    xValores(2) = count
    yValores(1) = promedio
    yValores(2) = promedio
    
    Set serie = cht.SeriesCollection.NewSeries
    With serie
        .Name = "Promedio"
        .XValues = xValores
        .Values = yValores
        .Format.Line.ForeColor.RGB = RGB(0, 0, 255)
        .Format.Line.DashStyle = msoLineDash
        .Format.Line.Weight = 2.5  ' ?? Línea más gruesa para mejor visibilidad
        .MarkerStyle = xlMarkerStyleNone
    End With
    
    ' Línea de promedio + 1s
    yValores(1) = promedio + desviacion
    yValores(2) = promedio + desviacion
    
    Set serie = cht.SeriesCollection.NewSeries
    With serie
        .Name = "Promedio + 1s"
        .XValues = xValores
        .Values = yValores
        .Format.Line.ForeColor.RGB = RGB(255, 165, 0)
        .Format.Line.DashStyle = msoLineDash
        .Format.Line.Weight = 2#   ' ?? Más gruesa
        .MarkerStyle = xlMarkerStyleNone
    End With
    
    ' Línea de promedio - 1s
    yValores(1) = promedio - desviacion
    yValores(2) = promedio - desviacion
    
    Set serie = cht.SeriesCollection.NewSeries
    With serie
        .Name = "Promedio - 1s"
        .XValues = xValores
        .Values = yValores
        .Format.Line.ForeColor.RGB = RGB(255, 165, 0)
        .Format.Line.DashStyle = msoLineDash
        .Format.Line.Weight = 2#   ' ?? Más gruesa
        .MarkerStyle = xlMarkerStyleNone
    End With
End Sub

' Función para crear gráfico de dispersión
Private Function CrearGraficoDispersion(ws As Worksheet, stats As EstadisticasColumna, topPosition As Long, leftColumn As Integer) As Long
    Dim cht As ChartObject
    Dim serie As Series
    Dim i As Long
    Dim xValores() As Variant
    Dim yValores() As Variant
    
    ' Preparar datos para el gráfico de dispersión
    ReDim xValores(1 To stats.count)
    ReDim yValores(1 To stats.count)
    
    For i = 1 To stats.count
        xValores(i) = i ' Eje X: secuencia
        yValores(i) = stats.valores(i - 1) ' Eje Y: valor
    Next i
    
    ' ?? GRÁFICO MÁS GRANDE: Mismo tamaño que el de control
    Set cht = ws.ChartObjects.Add(Left:=ws.Cells(topPosition, leftColumn).Left, _
                                  Top:=ws.Cells(topPosition, leftColumn).Top, _
                                  Width:=ANCHO_GRAFICO, _
                                  Height:=ALTO_GRAFICO)
    
    With cht.Chart
        .ChartType = xlXYScatter
        .HasTitle = True
        .chartTitle.Text = "Gráfico de Dispersión - " & stats.NombreColumna
        .chartTitle.Font.Size = 11  ' ?? Más grande
        .chartTitle.Font.Bold = True
        
        ' Configurar ejes
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Secuencia"
            .AxisTitle.Font.Size = 9  ' ?? Más grande
        End With
        
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Valor"
            .AxisTitle.Font.Size = 9  ' ?? Más grande
        End With
        
        ' Añadir serie de datos
        Set serie = .SeriesCollection.NewSeries
        With serie
            .Name = "Distribución"
            .XValues = xValores
            .Values = yValores
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 6  ' ?? Marcadores más grandes
            .MarkerForegroundColor = RGB(0, 100, 0)
            .MarkerBackgroundColor = RGB(0, 150, 0)
            .Format.Line.Visible = msoFalse
        End With
        
        ' Añadir línea de promedio
        AddLineaPromedioDispersion cht.Chart, stats.promedio, stats.minimo, stats.maximo
        
        ' Formato del gráfico
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(240, 240, 240)
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 9  ' ?? Más legible
    End With
    
    CrearGraficoDispersion = topPosition + (ALTO_GRAFICO / 15)
End Function

' Función para añadir línea de promedio al gráfico de dispersión
Private Sub AddLineaPromedioDispersion(cht As Chart, promedio As Double, minimo As Double, maximo As Double)
    Dim serie As Series
    Dim xValores() As Variant
    Dim yValores() As Variant
    
    ' Preparar datos para la línea de promedio
    ReDim xValores(1 To 2)
    ReDim yValores(1 To 2)
    
    xValores(1) = promedio
    xValores(2) = promedio
    yValores(1) = minimo - (maximo - minimo) * 0.1
    yValores(2) = maximo + (maximo - minimo) * 0.1
    
    Set serie = cht.SeriesCollection.NewSeries
    With serie
        .Name = "Promedio"
        .XValues = xValores
        .Values = yValores
        .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        .Format.Line.DashStyle = msoLineDash
        .Format.Line.Weight = 2.5  ' ?? Línea más gruesa
        .MarkerStyle = xlMarkerStyleNone
    End With
End Sub

' ============================================================================
' NUEVA FUNCIÓN: HISTOGRAMA CON CAMPANA DE GAUSS
' ============================================================================
Private Function CrearHistogramaConGauss(ws As Worksheet, stats As EstadisticasColumna, topPosition As Long, leftColumn As Integer) As Long
    Dim cht As ChartObject
    Dim serie As Series
    Dim i As Long, j As Long
    Dim numBins As Integer
    Dim binWidth As Double
    Dim minVal As Double, maxVal As Double
    
    ' Determinar número de bins (Regla de Sturges)
    numBins = Application.WorksheetFunction.RoundUp(1 + 3.322 * Application.WorksheetFunction.Log10(stats.count), 0)
    If numBins < 5 Then numBins = 5
    If numBins > 15 Then numBins = 15
    
    ' Calcular rango y ancho de bins
    minVal = stats.minimo
    maxVal = stats.maximo
    binWidth = (maxVal - minVal) / numBins
    
    ' Arrays para histograma
    Dim binLabels() As String
    Dim binCounts() As Long
    Dim binCenters() As Double
    ReDim binLabels(1 To numBins)
    ReDim binCounts(1 To numBins)
    ReDim binCenters(1 To numBins)
    
    ' Inicializar bins
    For i = 1 To numBins
        binCenters(i) = minVal + (i - 0.5) * binWidth
        binLabels(i) = Format(minVal + (i - 1) * binWidth, "0.0")
        binCounts(i) = 0
    Next i
    
    ' Contar frecuencias en cada bin
    For i = 0 To stats.count - 1
        Dim valor As Double
        valor = stats.valores(i)
        
        ' Determinar a qué bin pertenece
        Dim binIndex As Integer
        binIndex = Int((valor - minVal) / binWidth) + 1
        If binIndex > numBins Then binIndex = numBins
        If binIndex < 1 Then binIndex = 1
        
        binCounts(binIndex) = binCounts(binIndex) + 1
    Next i
    
    ' Crear gráfico
    Set cht = ws.ChartObjects.Add(Left:=ws.Cells(topPosition, leftColumn).Left, _
                                  Top:=ws.Cells(topPosition, leftColumn).Top, _
                                  Width:=ANCHO_GRAFICO, _
                                  Height:=ALTO_GRAFICO)
    
    With cht.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .chartTitle.Text = "Histograma con Distribución Normal - " & stats.NombreColumna
        .chartTitle.Font.Size = 11
        .chartTitle.Font.Bold = True
        
        ' Serie del histograma
        Set serie = .SeriesCollection.NewSeries
        With serie
            .Name = "Frecuencia"
            .XValues = binLabels
            .Values = binCounts
            .Format.Fill.ForeColor.RGB = RGB(237, 125, 49) ' Naranja
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            .Format.Line.Weight = 1.5
        End With
        
        ' Configurar ejes
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = stats.NombreColumna
            .AxisTitle.Font.Size = 9
        End With
        
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Frecuencia"
            .AxisTitle.Font.Size = 9
        End With
        
        ' Añadir curva de distribución normal
        Call AddCurvaGauss(cht.Chart, stats, binCenters, numBins, binWidth, stats.count)
        
        ' Formato del gráfico
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(240, 240, 240)
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 9
    End With
    
    CrearHistogramaConGauss = topPosition + (ALTO_GRAFICO / 15)
End Function

' Función para añadir curva de Gauss al histograma
Private Sub AddCurvaGauss(cht As Chart, stats As EstadisticasColumna, binCenters() As Double, numBins As Integer, binWidth As Double, totalCount As Long)
    Dim serie As Series
    Dim i As Integer
    Dim xGauss() As Double
    Dim yGauss() As Double
    Dim numPuntos As Integer
    
    ' Generar más puntos para una curva suave
    numPuntos = numBins * 5
    ReDim xGauss(1 To numPuntos)
    ReDim yGauss(1 To numPuntos)
    
    Dim minVal As Double, maxVal As Double
    minVal = stats.minimo
    maxVal = stats.maximo
    Dim paso As Double
    paso = (maxVal - minVal) / (numPuntos - 1)
    
    ' Calcular curva de Gauss
    Dim x As Double
    Dim gaussValue As Double
    Dim factor As Double
    
    ' Factor de escala para que la curva se ajuste al histograma
    factor = totalCount * binWidth
    
    For i = 1 To numPuntos
        x = minVal + (i - 1) * paso
        xGauss(i) = x
        
        ' Fórmula de distribución normal: (1/(sv(2p))) * e^(-(x-µ)²/(2s²))
        Dim exponente As Double
        exponente = -((x - stats.promedio) ^ 2) / (2 * stats.desviacionEstandar ^ 2)
        gaussValue = (1 / (stats.desviacionEstandar * Sqr(2 * 3.14159265358979))) * exp(exponente)
        
        ' Escalar para que coincida con el histograma
        yGauss(i) = gaussValue * factor
    Next i
    
    ' Añadir serie de curva (eje secundario)
    Set serie = cht.SeriesCollection.NewSeries
    With serie
        .Name = "Distribución Normal"
        .XValues = xGauss
        .Values = yGauss
        .ChartType = xlXYScatterLines
        .AxisGroup = xlSecondary
        .Format.Line.ForeColor.RGB = RGB(0, 0, 0) ' Negro
        .Format.Line.Weight = 2.5
        .MarkerStyle = xlMarkerStyleNone
    End With
    
    ' Configurar eje secundario
    With cht.Axes(xlValue, xlSecondary)
        .HasTitle = False
        .TickLabels.Font.Size = 8
    End With
End Sub

' Función auxiliar para verificar si hay suficientes datos para gráficos
Public Function HaySuficientesDatosParaGraficos(stats As EstadisticasColumna) As Boolean
    HaySuficientesDatosParaGraficos = (stats.count >= 3)
End Function
