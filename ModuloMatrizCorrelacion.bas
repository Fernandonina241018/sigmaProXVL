'================================================================================
' MÓDULO: ModuloMatrizCorrelacion
' PROPÓSITO: Generar matriz de correlación completa entre variables analizadas
' VERSIÓN: 1.0
' DEPENDENCIAS: Módulo Principal sigmaproxvl
' VALIDACIÓN: Métodos estadísticos según USP <1033> - Data Analysis
'================================================================================

Option Explicit
' ^-----------------------------------------------------------------------------
' | DECLARACIÓN OBLIGATORIA: Force explicit variable declaration
' | CUMPLIMIENTO: Buenas prácticas de programación para sistemas validados
' ------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
' TIPO DE DATO: MatrizCorrelacion
' PROPÓSITO: Almacenar estructura completa de matriz de correlación
' UTILIDAD: Permite cálculos intermedios y formato de salida
'--------------------------------------------------------------------------------
Type MatrizCorrelacion
    NombresVariables() As String    ' Array con nombres de variables
    Coeficientes() As Double        ' Matriz de coeficientes de correlación
    numVariables As Integer         ' Número de variables en la matriz
    FechaCalculo As Date            ' Timestamp del cálculo
    Metodo As String                ' Método utilizado (Pearson/Spearman)
End Type

'================================================================================
' FUNCIÓN: CrearMatrizCorrelacion
' PROPÓSITO: Función principal que coordina la creación de la matriz
' PARÁMETROS:
'   - stats(): Array de EstadisticasColumna con datos de análisis previo
'   - wb: Workbook donde generar la matriz
' RETORNO: Boolean (True=éxito, False=fallo)
' VALIDACIÓN: Verifica datos suficientes antes del cálculo
'================================================================================
Function CrearMatrizCorrelacion(stats() As EstadisticasColumna, wb As Workbook) As Boolean
    '----------------------------------------------------------------------------
    ' SECCIÓN: Declaración de Variables
    '----------------------------------------------------------------------------
    Dim matriz As MatrizCorrelacion
    Dim ws As Worksheet
    Dim nombreHoja As String
    Dim numeroHoja As Integer
    Dim i As Integer, j As Integer
    Dim numVariablesValidas As Integer
    Dim indicesValidos() As Integer
    Dim varCount As Integer
    
    On Error GoTo ErrorHandler
    
    '----------------------------------------------------------------------------
    ' SECCIÓN 1: Validación de Datos de Entrada
    '----------------------------------------------------------------------------
    ' Verificar que hay al menos 2 columnas con datos suficientes
    numVariablesValidas = 0
    ReDim indicesValidos(1 To UBound(stats))
    
    For i = 1 To UBound(stats)
        ' CRITERIO: Mínimo 3 datos por variable para cálculo de correlación
        If stats(i).count >= 3 Then
            numVariablesValidas = numVariablesValidas + 1
            indicesValidos(numVariablesValidas) = i
        End If
    Next i
    
    ' VALIDACIÓN: Requisito mínimo de variables
    If numVariablesValidas < 2 Then
        Debug.Print "Se requieren al menos 2 variables con 3 o más datos para generar " & _
               "la matriz de correlación." & vbCrLf & _
               "Variables válidas encontradas: " & numVariablesValidas, _
               vbExclamation, "Datos Insuficientes"
        CrearMatrizCorrelacion = False
        Exit Function
    End If
    
    '----------------------------------------------------------------------------
    ' SECCIÓN 2: Preparación de Estructura de Matriz
    '----------------------------------------------------------------------------
    With matriz
        .numVariables = numVariablesValidas
        .FechaCalculo = Now
        .Metodo = "Pearson"  ' Método estándar para datos paramétricos
        
        ' Dimensionar arrays
        ReDim .NombresVariables(1 To numVariablesValidas)
        ReDim .Coeficientes(1 To numVariablesValidas, 1 To numVariablesValidas)
        
        ' Almacenar nombres de variables válidas
        For i = 1 To numVariablesValidas
            .NombresVariables(i) = stats(indicesValidos(i)).NombreColumna
        Next i
    End With
    
    '----------------------------------------------------------------------------
    ' SECCIÓN 3: Cálculo de Coeficientes de Correlación
    '----------------------------------------------------------------------------
    If Not CalcularCoeficientesCorrelacion(stats, indicesValidos, numVariablesValidas, matriz) Then
        Debug.Print "Error en el cálculo de coeficientes de correlación.", vbCritical
        CrearMatrizCorrelacion = False
        Exit Function
    End If
    
    '----------------------------------------------------------------------------
    ' SECCIÓN 4: Creación de Hoja de Resultados
    '----------------------------------------------------------------------------
    If Not CrearHojaMatrizCorrelacion(matriz, wb) Then
        Debug.Print "Error al crear la hoja de matriz de correlación.", vbCritical
        CrearMatrizCorrelacion = False
        Exit Function
    End If
    
    CrearMatrizCorrelacion = True
    Exit Function

ErrorHandler:
    Debug.Print "Error inesperado en CrearMatrizCorrelacion: " & Err.Description, vbCritical
    CrearMatrizCorrelacion = False
End Function

'================================================================================
' FUNCIÓN: CalcularCoeficientesCorrelacion
' PROPÓSITO: Calcular coeficientes de correlación Pearson entre todas las variables
' PARÁMETROS:
'   - stats(): Array de EstadisticasColumna
'   - indicesValidos(): Array con índices de variables válidas
'   - numVariables: Número de variables a considerar
'   - matriz: MatrizCorrelacion (ByRef) para almacenar resultados
' RETORNO: Boolean (True=éxito, False=fallo)
' ALGORITMO: Coeficiente de correlación Pearson (r)
' FÓRMULA: r = S[(xi - x¯)(yi - ?)] / v[S(xi - x¯)² * S(yi - ?)²]
' REFERENCIA: USP <1033> - Correlation Analysis
'================================================================================
Function CalcularCoeficientesCorrelacion(stats() As EstadisticasColumna, _
                                        indicesValidos() As Integer, _
                                        numVariables As Integer, _
                                        ByRef matriz As MatrizCorrelacion) As Boolean
    '----------------------------------------------------------------------------
    ' SECCIÓN: Declaración de Variables
    '----------------------------------------------------------------------------
    Dim i As Integer, j As Integer, k As Integer
    Dim sumXY As Double, sumX As Double, sumY As Double
    Dim sumX2 As Double, sumY2 As Double
    Dim n As Long
    Dim x As Double, Y As Double
    Dim countPares As Long
    Dim valoresX() As Double, valoresY() As Double
    
    On Error GoTo ErrorHandler
    
    '----------------------------------------------------------------------------
    ' SECCIÓN: Cálculo de Matriz Completa
    '----------------------------------------------------------------------------
    For i = 1 To numVariables
        For j = 1 To numVariables
            If i = j Then
                ' DIAGONAL PRINCIPAL: Correlación consigo misma = 1
                matriz.Coeficientes(i, j) = 1
            Else
                '----------------------------------------------------------------
                ' SUBSECCIÓN: Preparación de Datos para Par Actual
                '----------------------------------------------------------------
                Dim idxI As Integer, idxJ As Integer
                idxI = indicesValidos(i)
                idxJ = indicesValidos(j)
                
                ' Obtener valores de ambas variables
                valoresX = stats(idxI).valores
                valoresY = stats(idxJ).valores
                
                '----------------------------------------------------------------
                ' SUBSECCIÓN: Alineación de Pares de Datos
                ' IMPORTANCIA: Correlación requiere pares alineados temporalmente
                '----------------------------------------------------------------
                n = Application.WorksheetFunction.Min(UBound(valoresX) + 1, _
                                                     UBound(valoresY) + 1)
                
                If n < 3 Then
                    matriz.Coeficientes(i, j) = 0
                Else
                    '------------------------------------------------------------
                    ' SUBSECCIÓN: Inicialización de Acumuladores
                    '------------------------------------------------------------
                    sumXY = 0
                    sumX = 0
                    sumY = 0
                    sumX2 = 0
                    sumY2 = 0
                    countPares = 0
                    
                    '------------------------------------------------------------
                    ' SUBSECCIÓN: Cálculo de Sumatorias
                    '------------------------------------------------------------
                    For k = 0 To n - 1
                        x = valoresX(k)
                        Y = valoresY(k)
                        
                        sumX = sumX + x
                        sumY = sumY + Y
                        sumXY = sumXY + (x * Y)
                        sumX2 = sumX2 + (x * x)
                        sumY2 = sumY2 + (Y * Y)
                        countPares = countPares + 1
                    Next k
                    
                    '------------------------------------------------------------
                    ' SUBSECCIÓN: Cálculo Final del Coeficiente
                    ' FÓRMULA: r = [nSXY - (SX)(SY)] / v{[nSX² - (SX)²][nSY² - (SY)²]}
                    '------------------------------------------------------------
                    If countPares >= 3 Then
                        Dim numerador As Double, denominador As Double
                        Dim denomX As Double, denomY As Double
                        
                        numerador = (countPares * sumXY) - (sumX * sumY)
                        denomX = (countPares * sumX2) - (sumX * sumX)
                        denomY = (countPares * sumY2) - (sumY * sumY)
                        
                        If denomX > 0 And denomY > 0 Then
                            denominador = Sqr(denomX * denomY)
                            If denominador <> 0 Then
                                matriz.Coeficientes(i, j) = numerador / denominador
                            Else
                                matriz.Coeficientes(i, j) = 0
                            End If
                        Else
                            matriz.Coeficientes(i, j) = 0
                        End If
                    Else
                        matriz.Coeficientes(i, j) = 0
                    End If
                End If
            End If
        Next j
    Next i
    
    CalcularCoeficientesCorrelacion = True
    Exit Function

ErrorHandler:
    Debug.Print "Error en CalcularCoeficientesCorrelacion: " & Err.Description, vbCritical
    CalcularCoeficientesCorrelacion = False
End Function

'================================================================================
' FUNCIÓN: CrearHojaMatrizCorrelacion
' PROPÓSITO: Crear hoja Excel con matriz de correlación formateada
' PARÁMETROS:
'   - matriz: MatrizCorrelacion con datos calculados
'   - wb: Workbook destino
' RETORNO: Boolean (True=éxito, False=fallo)
' FORMATO: Tabla profesional con coloración condicional
'================================================================================
Function CrearHojaMatrizCorrelacion(matriz As MatrizCorrelacion, wb As Workbook) As Boolean
    '----------------------------------------------------------------------------
    ' SECCIÓN: Declaración de Variables
    '----------------------------------------------------------------------------
    Dim ws As Worksheet
    Dim nombreHoja As String
    Dim numeroHoja As Integer
    Dim i As Integer, j As Integer
    Dim filaInicio As Integer, colInicio As Integer
    Dim rangoMatriz As Range
    Dim rangoDiagonal As Range
    
    On Error GoTo ErrorHandler
    
    '----------------------------------------------------------------------------
    ' SECCIÓN 1: Configuración Inicial de Hoja
    '----------------------------------------------------------------------------
    ' Obtener nombre único para hoja
    numeroHoja = ObtenerProximoNumeroHoja(wb, "Matriz Correlación")
    nombreHoja = "Matriz Correlación " & numeroHoja
    
    ' Crear nueva hoja
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.count))
    ws.Name = nombreHoja
    
    '----------------------------------------------------------------------------
    ' SECCIÓN 2: Encabezado y Metadatos
    '----------------------------------------------------------------------------
    With ws
        ' Encabezado Principal
        .Range("A1").Value = "MATRIZ DE CORRELACIÓN - ANÁLISIS MULTIVARIABLE"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1:G1").Merge
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Interior.color = RGB(200, 200, 200)
        
        ' Información del Análisis
        .Range("A2").Value = "Fecha de generación:"
        .Range("B2").Value = Format(matriz.FechaCalculo, "dd/mmm/yyyy hh:mm AM/PM")
        
        .Range("A3").Value = "Método estadístico:"
        .Range("B3").Value = matriz.Metodo & " (Coeficiente r)"
        
        .Range("A4").Value = "Número de variables:"
        .Range("B4").Value = matriz.numVariables
        
        .Range("A5").Value = "Interpretación:"
        .Range("B5").Value = "r ˜ ±1: Correlación fuerte | r ˜ 0: Sin correlación"
        .Range("B5").Font.Italic = True
        .Range("B5:G5").Merge
    End With
    
    '----------------------------------------------------------------------------
    ' SECCIÓN 3: Construcción de la Matriz
    '----------------------------------------------------------------------------
    filaInicio = 7
    colInicio = 2
    
    ' Escribir nombres de variables en filas y columnas
    With ws
        ' Nombres en columnas (horizontal)
        For i = 1 To matriz.numVariables
            .Cells(filaInicio, colInicio + i).Value = matriz.NombresVariables(i)
            .Cells(filaInicio, colInicio + i).Font.Bold = True
            .Cells(filaInicio, colInicio + i).Interior.color = RGB(220, 230, 241)
            .Cells(filaInicio, colInicio + i).HorizontalAlignment = xlCenter
        Next i
        
        ' Nombres en filas (vertical)
        For i = 1 To matriz.numVariables
            .Cells(filaInicio + i, colInicio - 1).Value = matriz.NombresVariables(i)
            .Cells(filaInicio + i, colInicio - 1).Font.Bold = True
            .Cells(filaInicio + i, colInicio - 1).Interior.color = RGB(220, 230, 241)
        Next i
        
        ' Escribir coeficientes de correlación
        For i = 1 To matriz.numVariables
            For j = 1 To matriz.numVariables
                .Cells(filaInicio + i, colInicio + j).Value = matriz.Coeficientes(i, j)
                .Cells(filaInicio + i, colInicio + j).NumberFormat = "0.0000"
            Next j
        Next i
    End With
    
    '----------------------------------------------------------------------------
    ' SECCIÓN 4: Formateo Avanzado de la Matriz
    '----------------------------------------------------------------------------
    Set rangoMatriz = ws.Range(ws.Cells(filaInicio + 1, colInicio), _
                              ws.Cells(filaInicio + matriz.numVariables, _
                              colInicio + matriz.numVariables))
    
    ' Aplicar bordes a toda la matriz
    With rangoMatriz
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' Coloración condicional para coeficientes
    AplicarColoracionCondicionalMatriz ws, rangoMatriz
    
    '----------------------------------------------------------------------------
    ' SECCIÓN 5: Ajustes Finales y Activación
    '----------------------------------------------------------------------------
    With ws
        ' Ajustar anchos de columna
        .Columns(colInicio - 1).ColumnWidth = 25
        For i = colInicio To colInicio + matriz.numVariables
            .Columns(i).ColumnWidth = 12
        Next i
        
        ' Autoajustar filas
        .Rows(filaInicio & ":" & filaInicio + matriz.numVariables).AutoFit
        
        ' Congelar paneles para fácil navegación
        .Cells(filaInicio + 1, colInicio).Select
        ActiveWindow.FreezePanes = True
        
        ' Activar la hoja
        .Activate
    End With
    
    CrearHojaMatrizCorrelacion = True
    Exit Function

ErrorHandler:
    Debug.Print "Error en CrearHojaMatrizCorrelacion: " & Err.Description, vbCritical
    CrearHojaMatrizCorrelacion = False
End Function

'================================================================================
' SUBRUTINA: AplicarColoracionCondicionalMatriz
' PROPÓSITO: Aplicar escala de colores para fácil interpretación de correlaciones
' ESCALA: Rojo (-1) ? Blanco (0) ? Verde (+1)
' INTERPRETACIÓN: Visualización inmediata de relaciones entre variables
'================================================================================
Sub AplicarColoracionCondicionalMatriz(ws As Worksheet, rango As Range)
    '----------------------------------------------------------------------------
    ' SECCIÓN: Configuración de Escala de Colores 3-Color
    '----------------------------------------------------------------------------
    With rango.FormatConditions.AddColorScale(ColorScaleType:=3)
        ' Color para valores bajos (correlación negativa fuerte)
        .ColorScaleCriteria(1).Type = xlConditionValueNumber
        .ColorScaleCriteria(1).Value = -1
        .ColorScaleCriteria(1).FormatColor.color = RGB(255, 0, 0)  ' Rojo
        
        ' Color para valores medios (sin correlación)
        .ColorScaleCriteria(2).Type = xlConditionValueNumber
        .ColorScaleCriteria(2).Value = 0
        .ColorScaleCriteria(2).FormatColor.color = RGB(255, 255, 255)  ' Blanco
        
        ' Color para valores altos (correlación positiva fuerte)
        .ColorScaleCriteria(3).Type = xlConditionValueNumber
        .ColorScaleCriteria(3).Value = 1
        .ColorScaleCriteria(3).FormatColor.color = RGB(0, 176, 80)  ' Verde
    End With
    
    '----------------------------------------------------------------------------
    ' SECCIÓN: Formato Adicional para Mejor Legibilidad
    '----------------------------------------------------------------------------
    With rango
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 9
    End With
End Sub

'================================================================================
' FUNCIÓN: ObtenerProximoNumeroHoja
' PROPÓSITO: Generar nombres de hojas únicos (réplica del módulo principal)
' PARÁMETROS:
'   - wb: Workbook objeto
'   - baseNombre: String con nombre base
' RETORNO: Integer con próximo número disponible
'================================================================================
Private Function ObtenerProximoNumeroHoja(wb As Workbook, baseNombre As String) As Integer
    Dim i As Integer
    i = 1
    While HojaExiste(wb, baseNombre & " " & i)
        i = i + 1
    Wend
    ObtenerProximoNumeroHoja = i
End Function

'================================================================================
' FUNCIÓN: HojaExiste
' PROPÓSITO: Verificar existencia de hoja en workbook
' PARÁMETROS:
'   - wb: Workbook objeto
'   - nombreHoja: String con nombre de hoja
' RETORNO: Boolean (True=existe, False=no existe)
'================================================================================
Private Function HojaExiste(wb As Workbook, nombreHoja As String) As Boolean
    On Error Resume Next
    HojaExiste = (wb.Sheets(nombreHoja).Name <> "")
    On Error GoTo 0
End Function

'================================================================================
' FUNCIÓN: EjecutarAnalisisCorrelacion
' PROPÓSITO: Función pública de entrada para ejecutar análisis de correlación
' PARÁMETROS:
'   - stats(): Array de EstadisticasColumna
'   - wb: Workbook objeto
'   - activarCorrelacion: Boolean desde checkbox del userform
' RETORNO: Boolean (True=análisis completado, False=omitido o error)
' INTEGRACIÓN: Llamada desde módulo principal condicionada al checkbox
'================================================================================
Public Function EjecutarAnalisisCorrelacion(stats() As EstadisticasColumna, _
                                           wb As Workbook, _
                                           activarCorrelacion As Boolean) As Boolean
    '----------------------------------------------------------------------------
    ' SECCIÓN: Validación de Condición de Ejecución
    '----------------------------------------------------------------------------
    If Not activarCorrelacion Then
        EjecutarAnalisisCorrelacion = False
        Exit Function
    End If
    
    '----------------------------------------------------------------------------
    ' SECCIÓN: Ejecución del Análisis
    '----------------------------------------------------------------------------
    If CrearMatrizCorrelacion(stats, wb) Then
        Debug.Print "Matriz de correlación generada exitosamente.", vbInformation
        EjecutarAnalisisCorrelacion = True
    Else
        EjecutarAnalisisCorrelacion = False
    End If
End Function

