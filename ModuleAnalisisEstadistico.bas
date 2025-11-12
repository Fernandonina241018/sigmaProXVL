'================================================================================
' MÓDULO: sigmaproxvl - Análisis Estadístico para Validación Farmacéutica
' PROPÓSITO: Realizar análisis estadístico completo con detección de outliers
'           y generación de reportes para auditorías regulatorias
' VERSIÓN: 2.0
' FECHA: [Fecha de Implementación]
' LICENCIA: Uso Interno - Validado para entornos GMP
'================================================================================

Option Explicit
' ^-----------------------------------------------------------------------------
' | DECLARACIÓN OBLIGATORIA: Force explicit variable declaration to prevent
' | runtime errors and improve code maintenance. Required for validated systems.
' | BENEFICIOS:
' | - Detecta errores de compilación tempranos
' | - Mejora la legibilidad del código
' | - Cumple con estándares de programación GMP
' ------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
' TIPO DE DATO: EstadisticasColumna
' PROPÓSITO: Almacenar resultados estadísticos completos por columna analizada
' ESTRUCTURA: Contiene parámetros estadísticos básicos, robustos y de outliers
' VALIDACIÓN: Estructura utilizada en procesos de validación metodológica
'--------------------------------------------------------------------------------
Type EstadisticasColumna
    '-------------------------------------------------------------------------
    ' SECCIÓN: Información Identificativa
    ' PROPÓSITO: Metadata para trazabilidad del análisis
    '-------------------------------------------------------------------------
    NombreColumna As String      ' Nombre descriptivo de la columna (ej: "Temperatura")
    columna As String            ' Letra de columna Excel (ej: "A", "B", "C")
    rango As String              ' Dirección del rango analizado (ej: "A1:A10")
    
    '-------------------------------------------------------------------------
    ' SECCIÓN: Estadísticos Descriptivos Básicos
    ' PROPÓSITO: Parámetros estadísticos fundamentales según Farmacopea USP
    ' REFERENCIA: Chapter <1033> USP - Data Analysis
    '-------------------------------------------------------------------------
    count As Long                ' Número de observaciones válidas (n)
    promedio As Double           ' Media aritmética (µ)
    desviacionEstandar As Double ' Desviación estándar muestral (s)
    RSD As Double                ' Coeficiente de variación (%RSD)
    maximo As Double             ' Valor máximo del conjunto
    minimo As Double             ' Valor mínimo del conjunto
    mediana As Double             ' Valor mínimo del conjunto
    varianza As Double           ' Valor mínimo del conjunto
    moda As Double               ' Valor mínimo del conjunto
    asimetria As Double          ' Valor mínimo del conjunto
    curtosis As Double
    
    '-------------------------------------------------------------------------
    ' SECCIÓN: Detección de Outliers - Método IQR
    ' PROPÓSITO: Identificar valores atípicos usando método de Tukey
    ' ALGORITMO: IQR * 1.5 (estándar industria farmacéutica)
    ' REFERENCIA: FDA Guidance - Outlier Detection
    '-------------------------------------------------------------------------
    Q1 As Double                 ' Primer cuartil (Percentil 25)
    Q3 As Double                 ' Tercer cuartil (Percentil 75)
    IQR As Double                ' Rango intercuartílico (Q3 - Q1)
    LimiteInferiorOutlier As Double ' Límite inferior: Q1 - 1.5*IQR
    LimiteSuperiorOutlier As Double ' Límite superior: Q3 + 1.5*IQR
    NumOutliers As Integer       ' Cantidad de outliers detectados
    Outliers() As Double         ' Array con valores outliers identificados
    
    '-------------------------------------------------------------------------
    ' SECCIÓN: Estadísticos Robustos
    ' PROPÓSITO: Cálculos excluyendo outliers para análisis conservativo
    ' IMPORTANCIA: Proporciona estimaciones más robustas en presencia de outliers
    '-------------------------------------------------------------------------
    MediaRobusta As Double       ' Media calculada excluyendo outliers
    DesvEstandarRobusta As Double ' Desviación estándar excluyendo outliers
    RSDrobusto As Double         ' %RSD calculado excluyendo outliers
    
    '-------------------------------------------------------------------------
    ' SECCIÓN: Datos Crudos para Gráficos
    ' PROPÓSITO: Almacenamiento de valores originales para visualización
    ' TRAZABILIDAD: Mantiene vinculación entre análisis y datos fuente
    '-------------------------------------------------------------------------
    valores() As Double          ' Array con valores originales para gráficos
End Type

'================================================================================
' FUNCIÓN: IsArrayEmpty
' PROPÓSITO: Verificar si un array de doubles está vacío o no inicializado
' PARÁMETROS:
'   - arr(): Array de doubles a verificar
' RETORNO: Boolean (True = array vacío, False = contiene datos)
' VALIDACIÓN: Método robusto para manejo seguro de arrays
'================================================================================
Function IsArrayEmpty(arr() As Double) As Boolean
    On Error Resume Next        ' Prevenir error si array no dimensionado
    IsArrayEmpty = (UBound(arr) < LBound(arr)) ' Array vacío si UBound < LBound
    On Error GoTo 0             ' Restaurar manejo normal de errores
End Function

'================================================================================
' FUNCIÓN: ObtenerProximoNumeroHoja
' PROPÓSITO: Generar nombres de hojas únicos con numeración secuencial
' PARÁMETROS:
'   - wb: Workbook objeto donde crear la hoja
'   - baseNombre: String con nombre base (ej: "Análisis Estadístico")
' RETORNO: Integer con próximo número disponible
' IMPORTANCIA: Previene sobrescritura de análisis previos
'================================================================================
Function ObtenerProximoNumeroHoja(wb As Workbook, baseNombre As String) As Integer
    Dim i As Integer
    i = 1
    ' Búsqueda incremental hasta encontrar número disponible
    While HojaExiste(wb, baseNombre & " " & i)
        i = i + 1
    Wend
    ObtenerProximoNumeroHoja = i
End Function

'================================================================================
' FUNCIÓN: HojaExiste
' PROPÓSITO: Verificar existencia de hoja en workbook especificado
' PARÁMETROS:
'   - wb: Workbook objeto a verificar
'   - nombreHoja: String con nombre de hoja a buscar
' RETORNO: Boolean (True = existe, False = no existe)
' MANEJO DE ERRORES: Usa On Error Resume Next para verificación segura
'================================================================================
Function HojaExiste(wb As Workbook, nombreHoja As String) As Boolean
    On Error Resume Next        ' Prevenir error si hoja no existe
    HojaExiste = (wb.Sheets(nombreHoja).Name <> "") ' Verificar nombre no vacío
    On Error GoTo 0             ' Restaurar manejo normal de errores
End Function

'================================================================================
' FUNCIÓN: OrdenarArray
' PROPÓSITO: Ordenar array de números usando algoritmo Bubble Sort
' PARÁMETROS:
'   - arr(): Array de doubles a ordenar
' RETORNO: Array de doubles ordenado ascendentemente
' ALGORITMO: Bubble Sort (suficiente para volúmenes de datos farmacéuticos)
' COMPLEJIDAD: O(n²) - adecuado para n < 1000 (casos típicos validación)
'================================================================================
Function OrdenarArray(arr() As Double) As Double()
    Dim i As Long, j As Long
    Dim temp As Double
    Dim resultado() As Double
    Dim n As Long
    
    ' Verificar si array tiene datos
    If Not IsArrayEmpty(arr) Then
        n = UBound(arr) - LBound(arr) + 1
        ReDim resultado(LBound(arr) To UBound(arr))
        
        ' Copiar array original (preservar datos fuente)
        For i = LBound(arr) To UBound(arr)
            resultado(i) = arr(i)
        Next i
        
        ' ALGORITMO: Bubble Sort
        ' PROPÓSITO: Ordenamiento estable para cálculo de percentiles
        For i = LBound(resultado) To UBound(resultado) - 1
            For j = i + 1 To UBound(resultado)
                If resultado(i) > resultado(j) Then
                    ' Intercambiar elementos
                    temp = resultado(i)
                    resultado(i) = resultado(j)
                    resultado(j) = temp
                End If
            Next j
        Next i
    Else
        ' Array vacío - retornar array con cero (evita errores downstream)
        ReDim resultado(0 To 0)
        resultado(0) = 0
    End If
    
    OrdenarArray = resultado
End Function

'================================================================================
' FUNCIÓN: CalcularPercentil
' PROPÓSITO: Calcular percentiles usando método de interpolación lineal
' PARÁMETROS:
'   - valores(): Array de doubles ordenado
'   - n: Long con número de elementos
'   - percentil: Double con percentil deseado (ej: 0.25 para Q1)
' RETORNO: Double con valor del percentil calculado
' MÉTODO: Interpolación lineal entre posiciones adyacentes
' REFERENCIA: NIST Handbook - Percentile Calculation Methods
'================================================================================
Function CalcularPercentil(valores() As Double, n As Long, percentil As Double) As Double
    Dim pos As Double
    Dim lowerIndex As Long, upperIndex As Long
    
    ' Validación: array vacío
    If n = 0 Then
        CalcularPercentil = 0
        Exit Function
    End If
    
    ' Cálculo de posición usando método h*(n-1)
    ' donde h = percentil (0-1)
    pos = (n - 1) * percentil
    lowerIndex = Int(pos)        ' Índice inferior
    upperIndex = lowerIndex + 1  ' Índice superior
    
    ' Manejo de límites del array
    If upperIndex >= n Then
        ' Caso: percentil en extremo superior
        CalcularPercentil = valores(n - 1)
    Else
        ' INTERPOLACIÓN LINEAL:
        ' P = valores[lowerIndex] + (pos - lowerIndex) *
        '     (valores[upperIndex] - valores[lowerIndex])
        CalcularPercentil = valores(lowerIndex) + (pos - lowerIndex) * _
                           (valores(upperIndex) - valores(lowerIndex))
    End If
End Function

'================================================================================
' SUBRUTINA: CalcularOutliersIQR
' PROPÓSITO: Detectar valores atípicos usando método IQR de Tukey
' PARÁMETROS:
'   - valores(): Array de doubles con datos a analizar
'   - stats: EstadisticasColumna (ByRef) para almacenar resultados
' ALGORITMO: IQR * 1.5 (estándar industria farmacéutica)
' VALIDACIÓN: Método recomendado por FDA para detección de outliers
'================================================================================
Sub CalcularOutliersIQR(valores() As Double, ByRef stats As EstadisticasColumna)
    Dim valoresOrdenados() As Double
    Dim i As Long, j As Long
    Dim esOutlier As Boolean
    
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
    
    ' REQUISITO MÍNIMO: 4 puntos para cálculo IQR válido
    If stats.count < 4 Then
        stats.NumOutliers = 0
        Exit Sub
    End If
    
    ' PASO 1: Ordenar valores para cálculo de cuartiles
    valoresOrdenados = OrdenarArray(valores)
    
    ' PASO 2: Calcular cuartiles Q1 (25%) y Q3 (75%)
    stats.Q1 = CalcularPercentil(valoresOrdenados, stats.count, 0.25)
    stats.Q3 = CalcularPercentil(valoresOrdenados, stats.count, 0.75)
    stats.IQR = stats.Q3 - stats.Q1  ' Rango intercuartílico
    
    ' PASO 3: Calcular límites de outliers
    ' FÓRMULA: Límite = Q ± 1.5 * IQR (estándar Tukey)
    stats.LimiteInferiorOutlier = stats.Q1 - 1.5 * stats.IQR
    stats.LimiteSuperiorOutlier = stats.Q3 + 1.5 * stats.IQR
    
    ' PASO 4: Identificar outliers
    stats.NumOutliers = 0
    For i = 0 To stats.count - 1
        ' CRITERIO: Valor fuera de [Q1-1.5*IQR, Q3+1.5*IQR]
        If valores(i) < stats.LimiteInferiorOutlier Or valores(i) > stats.LimiteSuperiorOutlier Then
            ' Redimensionar array de outliers dinámicamente
            If stats.NumOutliers = 0 Then
                ReDim stats.Outliers(0 To 0)
            Else
                ReDim Preserve stats.Outliers(0 To stats.NumOutliers)
            End If
            ' Almacenar valor outlier
            stats.Outliers(stats.NumOutliers) = valores(i)
            stats.NumOutliers = stats.NumOutliers + 1
        End If
    Next i
    
    ' PASO 5: Calcular estadísticas robustas (excluyendo outliers)
    If stats.NumOutliers > 0 Then
        CalcularEstadisticasRobustas valores, stats
    Else
        ' Si no hay outliers, estadísticas robustas = estadísticas normales
        stats.MediaRobusta = stats.promedio
        stats.DesvEstandarRobusta = stats.desviacionEstandar
        If stats.MediaRobusta <> 0 Then
            stats.RSDrobusto = (stats.DesvEstandarRobusta / stats.MediaRobusta) * 100
        Else
            stats.RSDrobusto = 0
        End If
    End If
    
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
End Sub

'================================================================================
' SUBRUTINA: CalcularEstadisticasRobustas
' PROPÓSITO: Calcular estadísticas excluyendo valores atípicos
' PARÁMETROS:
'   - valores(): Array de doubles con datos originales
'   - stats: EstadisticasColumna (ByRef) para almacenar resultados
' IMPORTANCIA: Proporciona estimaciones conservativas para toma de decisiones
'================================================================================
Sub CalcularEstadisticasRobustas(valores() As Double, ByRef stats As EstadisticasColumna)
    Dim suma As Double, sumaCuadrados As Double
    Dim countRobusto As Long, i As Long, j As Long
    Dim esOutlier As Boolean
    
    Dim calcMode As XlCalculation
    calcMode = Application.Calculation
    
    ' DESACTIVAR durante procesamiento
    With Application
        .ScreenUpdating = False             ' No actualizar pantalla (CRÍTICO)
        .Calculation = xlCalculationManual  ' Desactivar cálculos automáticos
        .EnableEvents = False               ' Desactivar eventos
        .DisplayStatusBar = False           ' Ocultar barra de estado
    End With
    
    On Error GoTo Cleanup
    
    ' INICIALIZACIÓN: Resetear acumuladores
    suma = 0
    sumaCuadrados = 0
    countRobusto = 0
    
    ' PASO 1: Calcular suma excluyendo outliers
    For i = 0 To stats.count - 1
        esOutlier = False
        
        ' Verificar si el valor actual es outlier
        If stats.NumOutliers > 0 Then
            For j = 0 To stats.NumOutliers - 1
                If valores(i) = stats.Outliers(j) Then
                    esOutlier = True
                    Exit For
                End If
            Next j
        End If
        
        ' Si no es outlier, incluir en cálculos
        If Not esOutlier Then
            suma = suma + valores(i)
            countRobusto = countRobusto + 1
        End If
    Next i
    
    ' PASO 2: Calcular media robusta
    If countRobusto > 0 Then
        stats.MediaRobusta = suma / countRobusto
        
        ' PASO 3: Calcular suma de cuadrados para desviación estándar
        For i = 0 To stats.count - 1
            esOutlier = False
            If stats.NumOutliers > 0 Then
                For j = 0 To stats.NumOutliers - 1
                    If valores(i) = stats.Outliers(j) Then
                        esOutlier = True
                        Exit For
                    End If
                Next j
            End If
            
            ' Incluir en suma de cuadrados si no es outlier
            If Not esOutlier Then
                sumaCuadrados = sumaCuadrados + (valores(i) - stats.MediaRobusta) ^ 2
            End If
        Next i
        
        ' PASO 4: Calcular desviación estándar robusta
        If countRobusto > 1 Then
            ' FÓRMULA: Desviación estándar muestral (n-1)
            stats.DesvEstandarRobusta = Sqr(sumaCuadrados / (countRobusto - 1))
        Else
            stats.DesvEstandarRobusta = 0
        End If
        
        ' PASO 5: Calcular %RSD robusto
        If stats.MediaRobusta <> 0 Then
            stats.RSDrobusto = (stats.DesvEstandarRobusta / stats.MediaRobusta) * 100
        Else
            stats.RSDrobusto = 0
        End If
    Else
        ' CASO EXTREMO: Todos los datos son outliers
        stats.MediaRobusta = 0
        stats.DesvEstandarRobusta = 0
        stats.RSDrobusto = 0
    End If
    
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
    
End Sub

'================================================================================
' FUNCIÓN: AnalizarColumna
' PROPÓSITO: Función principal que realiza análisis estadístico completo
' PARÁMETROS:
'   - rangoColumna: Range objeto Excel con datos a analizar
' RETORNO: EstadisticasColumna con resultados completos
' ALCANCE: Análisis completo desde extracción hasta cálculo estadístico
'================================================================================
Function AnalizarColumna(rangoColumna As Range) As EstadisticasColumna
    Dim resultados As EstadisticasColumna
    Dim celda As Range
    Dim suma As Double
    Dim sumaCuadrados As Double
    Dim valores() As Double
    Dim i As Long
    
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
    
    '----------------------------------------------------------------------------
    ' SECCIÓN 1: INICIALIZACIÓN
    '----------------------------------------------------------------------------
    resultados.count = 0
    suma = 0
    sumaCuadrados = 0
    ' Inicializar con valores extremos para búsqueda de min/max
    resultados.maximo = -1.79769313486231E+308  ' Mínimo valor Double negativo
    resultados.minimo = 1.79769313486231E+308   ' Máximo valor Double positivo
    resultados.rango = rangoColumna.Address     ' Guardar referencia de rango
    
    '----------------------------------------------------------------------------
    ' SECCIÓN 2: EXTRACCIÓN DE METADATA
    '----------------------------------------------------------------------------
    ' Obtener letra de columna de forma robusta
    On Error Resume Next
    resultados.columna = rangoColumna.Cells(1, 1).EntireColumn.Address(0, 0)
    resultados.columna = Replace(resultados.columna, ":", "") ' Limpiar formato
    On Error GoTo 0
    
    ' Obtener nombre descriptivo (asume primera fila = encabezado)
    On Error Resume Next
    resultados.NombreColumna = rangoColumna.Cells(1, 1).Value
    If resultados.NombreColumna = "" Then
        resultados.NombreColumna = "Columna " & resultados.columna ' Nombre por defecto
    End If
    On Error GoTo 0
    
    '----------------------------------------------------------------------------
    ' SECCIÓN 3: EXTRACCIÓN Y VALIDACIÓN DE DATOS
    '----------------------------------------------------------------------------
    For Each celda In rangoColumna.Cells
        ' EXCLUSIÓN: Saltar primera fila (asumida como encabezado)
        If celda.Row > rangoColumna.Cells(1, 1).Row Then
            ' VALIDACIÓN: Solo datos numéricos no vacíos
            If IsNumeric(celda.Value) And celda.Value <> "" Then
                ReDim Preserve valores(resultados.count)
                valores(resultados.count) = celda.Value
                suma = suma + celda.Value
                resultados.count = resultados.count + 1
            End If
        End If
    Next celda
    
    '----------------------------------------------------------------------------
    ' SECCIÓN 4: PRESERVACIÓN DE DATOS ORIGINALES
    '----------------------------------------------------------------------------
    If resultados.count > 0 Then
        ReDim resultados.valores(0 To resultados.count - 1)
        For i = 0 To resultados.count - 1
            resultados.valores(i) = valores(i) ' Copia para trazabilidad
        Next i
    End If
    
    '----------------------------------------------------------------------------
    ' SECCIÓN 5: CÁLCULOS ESTADÍSTICOS BÁSICOS
    '----------------------------------------------------------------------------
    If resultados.count > 0 Then
        ' CÁLCULO: Media aritmética
        resultados.promedio = suma / resultados.count
        
        ' CÁLCULO: Suma de cuadrados para desviación estándar
        For i = 0 To resultados.count - 1
            sumaCuadrados = sumaCuadrados + (valores(i) - resultados.promedio) ^ 2
        Next i
        
        ' CÁLCULO: Desviación estándar muestral (n-1)
        If resultados.count > 1 Then
            resultados.desviacionEstandar = Sqr(sumaCuadrados / (resultados.count - 1))
        Else
            resultados.desviacionEstandar = 0
        End If
        
        ' CÁLCULO: Coeficiente de variación (%RSD)
        If resultados.promedio <> 0 Then
            resultados.RSD = (resultados.desviacionEstandar / resultados.promedio) * 100
        Else
            resultados.RSD = 0
        End If
        
        If resultados.promedio <> 0 Then
            resultados.varianza = CalcularVarianza(valores)
        Else
            resultados.varianza = 0
        End If
        
        If resultados.promedio <> 0 Then
            resultados.mediana = CalcularMediana(valores)
        Else
            resultados.mediana = 0
        End If
        
        If resultados.promedio <> 0 Then
            resultados.moda = CalcularModa(valores)
        Else
            resultados.moda = 0
        End If
        
        If resultados.promedio <> 0 Then
            resultados.asimetria = CalcularAsimetria(valores)
        Else
            resultados.asimetria = 0
        End If
        
        If resultados.promedio <> 0 Then
            resultados.curtosis = CalcularCurtosis(valores)
        Else
            resultados.curtosis = 0
        End If
        
        ' CÁLCULO: Valores máximo y mínimo
        For i = 0 To resultados.count - 1
            If valores(i) > resultados.maximo Then resultados.maximo = valores(i)
            If valores(i) < resultados.minimo Then resultados.minimo = valores(i)
        Next i
        
        '----------------------------------------------------------------------------
        ' SECCIÓN 6: ANÁLISIS DE OUTLIERS Y ESTADÍSTICAS ROBUSTAS
        '----------------------------------------------------------------------------
        CalcularOutliersIQR valores, resultados
    Else
        ' CASO: Sin datos válidos - inicializar campos
        resultados.NumOutliers = 0
        resultados.MediaRobusta = 0
        resultados.DesvEstandarRobusta = 0
        resultados.RSDrobusto = 0
    End If
    
    AnalizarColumna = resultados
    
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
    
End Function

'================================================================================
' FUNCIÓN: ArrayToString
' PROPÓSITO: Convertir array de doubles a string formateado para display
' PARÁMETROS:
'   - arr(): Array de doubles a convertir
' RETORNO: String con valores formateados separados por comas
' FORMATO: "0.0000" (4 decimales para precisión en reportes)
'================================================================================
Function ArrayToString(arr() As Double) As String
    Dim i As Long
    Dim result As String
    result = ""
    
    If Not IsArrayEmpty(arr) Then
        For i = LBound(arr) To UBound(arr)
            If result <> "" Then result = result & ", "
            result = result & Format(arr(i), "0.0000") ' Formato de 4 decimales
        Next i
    End If
    
    ArrayToString = result
End Function

Public Sub EjecutarAnalisisCapacidad(stats() As EstadisticasColumna, wb As Workbook)
    Dim datosCapacidad As DatosCapacidadProceso
    Dim parametros As ParametrosCalculoCapacidad
    Dim resultados As ResultadosCapacidadProceso
    
    Dim calcMode As XlCalculation
    calcMode = Application.Calculation
    
    ' DESACTIVAR durante procesamiento
    With Application
        .ScreenUpdating = False             ' No actualizar pantalla (CRÍTICO)
        .Calculation = xlCalculationManual  ' Desactivar cálculos automáticos
        .EnableEvents = False               ' Desactivar eventos
        .DisplayStatusBar = False           ' Ocultar barra de estado
    End With
    
    On Error GoTo Cleanup
    
    ' Configurar datos desde el UserForm
    datosCapacidad = ObtenerDatosCapacidadDesdeFormulario()
    parametros = ObtenerParametrosCalculoDesdeFormulario()
    
    ' Ejecutar análisis para cada columna con datos suficientes
    Dim col As Integer
    For col = 1 To UBound(stats)
        If stats(col).count >= 10 Then ' Mínimo 10 datos
            datosCapacidad.valores = stats(col).valores
            datosCapacidad.NombreProceso = stats(col).NombreColumna
            
            resultados = AnalizarCapacidadProceso(datosCapacidad, parametros)
            'Call GenerarReporteCapacidad(wb, resultados, col)
        End If
    Next col
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
End Sub
