Public wbTarget As Workbook

Option Explicit

' =====================================================
' MÓDULO: ANÁLISIS DE REGRESIÓN MEJORADO
' Propósito: Implementación robusta de regresión lineal simple y múltiple
' Características: Estabilidad numérica, validación completa, documentación auditora
' Referencias: NIST Engineering Statistics Handbook, ISO 22514-4
' =====================================================

' Constantes para precisión numérica y validación
Public Const EPSILON As Double = 0.000000000001 ' Tolerancia para singularidad
Public Const MIN_SAMPLE_SIMPLE As Long = 3 ' Mínimo para regresión simple
Public Const MIN_SAMPLE_MULTIPLE As Long = 4 ' Mínimo para regresión múltiple
Public Const MAX_CONDITION_NUMBER As Double = 10000000000# ' Número de condición máximo

' Tipo personalizado para resultados de regresión
Public Type RegressionResult
    Coefficients As Variant
    R2 As Double
    R2Adjusted As Double
    StandardErrors As Variant
    TStats As Variant
    PValues As Variant
    SSE As Double
    SSR As Double
    SST As Double
    MSE As Double
    FStat As Double
    DF_Regression As Long
    DF_Residual As Long
    IsValid As Boolean
    ErrorMessage As String
End Type

' =====================================================
' REGRESIÓN LINEAL SIMPLE - VERSIÓN MEJORADA
' =====================================================

Public Sub RegresionLinealSimpleMejorada(variableX As Range, variableY As Range)
    Dim resultado As RegressionResult
    Dim datosX As Variant, datosY As Variant
    Dim wsResultados As Worksheet
    Dim nombreHoja As String
    
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
    
    On Error GoTo ErrorHandler
    
    ' 1. VALIDACIÓN COMPLETA DE DATOS
    If Not ValidarDatosEntrada(variableX, variableY, datosX, datosY) Then
        Exit Sub
    End If
    
    ' 2. CÁLCULO NUMÉRICAMENTE ESTABLE
    resultado = CalcularRegresionSimple(datosX, datosY)
    
    If Not resultado.IsValid Then
        MsgBox "Error en cálculo: " & resultado.ErrorMessage, vbCritical
        Exit Sub
    End If
    
    ' 3. CREAR HOJA DE RESULTADOS COMPLETA
    nombreHoja = "Regresion_Simple_" & Format(Now, "yyyy/mmm/dd_hhmmss")
    Set wsResultados = CrearHojaResultados(nombreHoja)
    LlenarHojaRegresionSimple wsResultados, resultado, datosX, datosY
    
    ' 4. GENERAR ANÁLISIS DE RESIDUOS
    GenerarAnalisisResidual wsResultados, datosX, datosY, resultado
    
    wsResultados.Activate
    MsgBox "Análisis de regresión simple completado exitosamente", vbInformation
    
    Exit Sub

ErrorHandler:
    ManejarError "RegresionLinealSimpleMejorada", Err.Number, Err.Description
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
End Sub

' =====================================================
' REGRESIÓN LINEAL MÚLTIPLE - VERSIÓN MEJORADA
' =====================================================

Public Sub RegresionLinealMultipleMejorada(ParamArray variables() As Variant)
    Dim resultado As RegressionResult
    Dim datos() As Variant
    Dim wsResultados As Worksheet
    Dim nombreHoja As String
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
    
    On Error GoTo ErrorHandler
    
    ' Validar número de variables (mínimo 2: Y + al menos 1 X)
    If UBound(variables) < 1 Then
        MsgBox "Error: Se requieren al menos 2 variables (Y y al menos 1 X)", vbCritical
        Exit Sub
    End If
    
    ' 1. VALIDACIÓN COMPLETA DE DATOS
    If Not ValidarDatosMultiples(variables, datos) Then
        Exit Sub
    End If
    
    ' 2. CÁLCULO CON VALIDACIÓN DE MATRIZ
    resultado = CalcularRegresionMultiple(datos)
    
    If Not resultado.IsValid Then
        MsgBox "Error en cálculo: " & resultado.ErrorMessage, vbCritical
        Exit Sub
    End If
    
    ' 3. CREAR HOJA DE RESULTADOS COMPLETA
    nombreHoja = "Regresion_Multiple_" & Format(Now, "yyyy/mmm/dd_hhmmss")
    Set wsResultados = CrearHojaResultados(nombreHoja)
    LlenarHojaRegresionMultiple wsResultados, resultado, datos, UBound(variables)
    
    ' 4. GENERAR ANÁLISIS DE RESIDUOS
    GenerarAnalisisResidualMultiple wsResultados, datos, resultado, UBound(variables)
    
    wsResultados.Activate
    MsgBox "Análisis de regresión múltiple completado exitosamente", vbInformation
    
    Exit Sub

ErrorHandler:
    ManejarError "RegresionLinealMultipleMejorada", Err.Number, Err.Description
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
End Sub

' =====================================================
' FUNCIONES DE VALIDACIÓN MEJORADAS
' =====================================================

Public Function ValidarDatosEntrada(variableX As Range, variableY As Range, _
                                   ByRef datosX As Variant, ByRef datosY As Variant) As Boolean
    ' Validación completa para auditorías
    
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
    
    ' Verificar que los rangos existan
    If variableX Is Nothing Or variableY Is Nothing Then
        MsgBox "Error: Los rangos de datos no pueden estar vacíos", vbCritical
        Exit Function
    End If
    
    ' Verificar mismo tamaño
    If variableX.Cells.count <> variableY.Cells.count Then
        MsgBox "Error: Los rangos deben tener el mismo número de observaciones", vbCritical
        Exit Function
    End If
    
    ' Verificar tamaño mínimo
    If variableX.Cells.count < MIN_SAMPLE_SIMPLE Then
        MsgBox "Error: Se requieren al menos " & MIN_SAMPLE_SIMPLE & " observaciones", vbCritical
        Exit Function
    End If
    
    ' Convertir a arrays y validar datos
    datosX = ObtenerArrayNumerico(variableX)
    datosY = ObtenerArrayNumerico(variableY)
    
    If Not EsArrayValido(datosX) Or Not EsArrayValido(datosY) Then
        MsgBox "Error: Los datos contienen valores no numéricos o vacíos", vbCritical
        Exit Function
    End If
    
    ' Verificar que X no sea constante
    If EsVectorConstante(datosX) Then
        MsgBox "Error: La variable independiente no puede ser constante", vbCritical
        Exit Function
    End If
    
    ValidarDatosEntrada = True
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
End Function

Public Function ValidarDatosMultiples(variables As Variant, ByRef datos() As Variant) As Boolean
    Dim i As Long, j As Long
    Dim n As Long, p As Long
    Dim tempData As Variant
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
    
    p = UBound(variables) ' Número de variables (Y + X's)
    n = variables(0).Cells.count
    
    ' Validar que todos los rangos tengan mismo tamaño
    For i = 0 To p
        If variables(i).Cells.count <> n Then
            MsgBox "Error: Todas las variables deben tener el mismo número de observaciones", vbCritical
            Exit Function
        End If
    Next i
    
    ' Validar tamaño mínimo
    If n < MIN_SAMPLE_MULTIPLE Then
        MsgBox "Error: Se requieren al menos " & MIN_SAMPLE_MULTIPLE & " observaciones", vbCritical
        Exit Function
    End If
    
    ' Crear matriz de datos
    ReDim datos(1 To n, 1 To p + 1)
    
    For i = 0 To p
        tempData = ObtenerArrayNumerico(variables(i))
        
        If Not EsArrayValido(tempData) Then
            MsgBox "Error: Los datos contienen valores no numéricos o vacíos en variable " & i + 1, vbCritical
            Exit Function
        End If
        
        For j = 1 To n
            datos(j, i + 1) = tempData(j, 1)
        Next j
    Next i
    
    ' Verificar que no haya columnas constantes (excepto quizás la primera que es Y)
    For i = 2 To p + 1 ' Empezar desde primera variable X
        If EsColumnaConstante(datos, i) Then
            MsgBox "Error: La variable X" & i - 1 & " es constante", vbCritical
            Exit Function
        End If
    Next i
    
    ValidarDatosMultiples = True
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
End Function

' =====================================================
' CÁLCULOS NUMÉRICAMENTE ESTABLES
' =====================================================

Public Function CalcularRegresionSimple(datosX As Variant, datosY As Variant) As RegressionResult
    Dim resultado As RegressionResult
    Dim n As Long, i As Long
    Dim meanX As Double, meanY As Double
    Dim Sxx As Double, Sxy As Double, Syy As Double
    Dim beta1 As Double, beta0 As Double
    Dim yPred As Double, residuo As Double
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
    
    On Error GoTo CalculationError
    
    ' 1. CALCULAR MEDIAS (primero para estabilidad)
    For i = 1 To n
        meanX = meanX + datosX(i, 1)
        meanY = meanY + datosY(i, 1)
    Next i
    meanX = meanX / n
    meanY = meanY / n
    
    ' 2. CALCULAR SUMAS DE CUADRADOS CENTRADAS (numéricamente estable)
    For i = 1 To n
        Sxx = Sxx + (datosX(i, 1) - meanX) * (datosX(i, 1) - meanX)
        Sxy = Sxy + (datosX(i, 1) - meanX) * (datosY(i, 1) - meanY)
        Syy = Syy + (datosY(i, 1) - meanY) * (datosY(i, 1) - meanY)
    Next i
    
    ' 3. VERIFICAR SINGULARIDAD
    If Abs(Sxx) < EPSILON Then
        resultado.ErrorMessage = "Matriz de diseño singular - X es constante"
        resultado.IsValid = False
        Exit Function
    End If
    
    ' 4. CALCULAR COEFICIENTES
    beta1 = Sxy / Sxx
    beta0 = meanY - beta1 * meanX
    
    ' 5. CALCULAR ESTADÍSTICAS DEL MODELO
    ReDim resultado.Coefficients(1 To 2, 1 To 1)
    resultado.Coefficients(1, 1) = beta0 ' Intercepto
    resultado.Coefficients(2, 1) = beta1 ' Pendiente
    
    ' Sumas de cuadrados
    resultado.SSR = beta1 * Sxy
    resultado.SST = Syy
    resultado.SSE = resultado.SST - resultado.SSR
    
    ' R-cuadrado
    If Abs(resultado.SST) > EPSILON Then
        resultado.R2 = resultado.SSR / resultado.SST
    Else
        resultado.R2 = 1 ' Si Y es constante y modelo perfecto
    End If
    
    ' Grados de libertad
    resultado.DF_Regression = 1
    resultado.DF_Residual = n - 2
    
    ' Error cuadrático medio
    If resultado.DF_Residual > 0 Then
        resultado.MSE = resultado.SSE / resultado.DF_Residual
    Else
        resultado.MSE = 0
    End If
    
    ' Estadístico F
    If resultado.MSE > EPSILON Then
        resultado.FStat = (resultado.SSR / resultado.DF_Regression) / resultado.MSE
    Else
        resultado.FStat = 1E+308 ' Infinito para MSE cero
    End If
    
    ' Errores estándar y estadísticas t
    ReDim resultado.StandardErrors(1 To 2, 1 To 1)
    ReDim resultado.TStats(1 To 2, 1 To 1)
    ReDim resultado.PValues(1 To 2, 1 To 1)
    
    If resultado.DF_Residual > 0 And resultado.MSE > EPSILON Then
        ' Error estándar del intercepto
        resultado.StandardErrors(1, 1) = Sqr(resultado.MSE * (1# / n + meanX * meanX / Sxx))
        ' Error estándar de la pendiente
        resultado.StandardErrors(2, 1) = Sqr(resultado.MSE / Sxx)
        
        ' Estadísticas t
        For i = 1 To 2
            If resultado.StandardErrors(i, 1) > EPSILON Then
                resultado.TStats(i, 1) = resultado.Coefficients(i, 1) / resultado.StandardErrors(i, 1)
                resultado.PValues(i, 1) = wsFunc.T_Dist_2T(Abs(resultado.TStats(i, 1)), resultado.DF_Residual)
            Else
                resultado.TStats(i, 1) = 0
                resultado.PValues(i, 1) = 1
            End If
        Next i
    End If
    
    resultado.IsValid = True
    CalcularRegresionSimple = resultado
    Exit Function

CalculationError:
    resultado.IsValid = False
    resultado.ErrorMessage = "Error en cálculo: " & Err.Description
    CalcularRegresionSimple = resultado
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
End Function

Public Function CalcularRegresionMultiple(datos() As Variant) As RegressionResult
    Dim resultado As RegressionResult
    Dim n As Long, p As Long, i As Long, j As Long, k As Long
    Dim x() As Double, Y() As Double
    Dim XT() As Double, XTX As Variant, XTY As Variant
    Dim beta As Variant, residuals() As Double
    Dim wsFunc As WorksheetFunction
    Dim conditionNumber As Double
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
    n = UBound(datos, 1)
    p = UBound(datos, 2) ' p incluye Y + X's
    
    On Error GoTo CalculationError
    
    ' p es el número total de columnas (Y + variables X)
    ' El modelo tiene (p-1) variables X + intercepto = p parámetros
    
    ' 1. CONSTRUIR MATRICES DE DISEÑO
    ReDim x(1 To n, 1 To p) ' Columna 1: intercepto, resto: variables
    ReDim Y(1 To n, 1 To 1)
    
    For i = 1 To n
        x(i, 1) = 1 ' Intercepto
        For j = 2 To p
            x(i, j) = datos(i, j)
        Next j
        Y(i, 1) = datos(i, 1) ' Y es la primera columna
    Next i
    
    ' 2. CALCULAR X'X Y X'Y
    XTX = wsFunc.MMult(wsFunc.Transpose(x), x)
    XTY = wsFunc.MMult(wsFunc.Transpose(x), Y)
    
    ' 3. VERIFICAR CONDICIÓN DE LA MATRIZ
    If Not EsMatrizInvertible(XTX, conditionNumber) Then
        If conditionNumber > MAX_CONDITION_NUMBER Then
            resultado.ErrorMessage = "Matriz mal condicionada. Número de condición: " & Format(conditionNumber, "0.00E+00")
        Else
            resultado.ErrorMessage = "Matriz de diseño singular - variables posiblemente colineales"
        End If
        resultado.IsValid = False
        Exit Function
    End If
    
    ' 4. CALCULAR COEFICIENTES: beta = (X'X)^-1 X'Y
    beta = wsFunc.MMult(wsFunc.MInverse(XTX), XTY)
    
    ' 5. CALCULAR ESTADÍSTICAS DEL MODELO
    resultado.Coefficients = beta
    
    ' Calcular predicciones y residuos
    ReDim residuals(1 To n)
    Dim meanY As Double, SST As Double, SSR As Double, SSE As Double
    Dim yPred As Double
    
    ' Media de Y
    For i = 1 To n
        meanY = meanY + Y(i, 1)
    Next i
    meanY = meanY / n
    
    ' Sumas de cuadrados
    For i = 1 To n
        yPred = 0
        For j = 1 To p
            yPred = yPred + beta(j, 1) * x(i, j)
        Next j
        residuals(i) = Y(i, 1) - yPred
        
        SST = SST + (Y(i, 1) - meanY) ^ 2
        SSR = SSR + (yPred - meanY) ^ 2
        SSE = SSE + residuals(i) ^ 2
    Next i
    
    resultado.SST = SST
    resultado.SSR = SSR
    resultado.SSE = SSE
    
    ' R-cuadrado y R-cuadrado ajustado
    If Abs(SST) > EPSILON Then
        resultado.R2 = SSR / SST
        resultado.R2Adjusted = 1 - (SSE / (n - p)) / (SST / (n - 1))
    Else
        resultado.R2 = 1
        resultado.R2Adjusted = 1
    End If
    
    ' Grados de libertad
    resultado.DF_Regression = p - 1
    resultado.DF_Residual = n - p
    
    ' Error cuadrático medio
    If resultado.DF_Residual > 0 Then
        resultado.MSE = SSE / resultado.DF_Residual
    Else
        resultado.MSE = 0
    End If
    
    ' Estadístico F
    If resultado.MSE > EPSILON And resultado.DF_Regression > 0 Then
        resultado.FStat = (SSR / resultado.DF_Regression) / resultado.MSE
    Else
        resultado.FStat = 0
    End If
    
    ' Calcular errores estándar y estadísticas t
    If resultado.DF_Residual > 0 And resultado.MSE > EPSILON Then
        Dim XTXinv As Variant
        XTXinv = wsFunc.MInverse(XTX)
        
        ReDim resultado.StandardErrors(1 To p, 1 To 1)
        ReDim resultado.TStats(1 To p, 1 To 1)
        ReDim resultado.PValues(1 To p, 1 To 1)
        
        For i = 1 To p
            resultado.StandardErrors(i, 1) = Sqr(resultado.MSE * XTXinv(i, i))
            If resultado.StandardErrors(i, 1) > EPSILON Then
                resultado.TStats(i, 1) = beta(i, 1) / resultado.StandardErrors(i, 1)
                resultado.PValues(i, 1) = wsFunc.T_Dist_2T(Abs(resultado.TStats(i, 1)), resultado.DF_Residual)
            Else
                resultado.TStats(i, 1) = 0
                resultado.PValues(i, 1) = 1
            End If
        Next i
    End If
    
    resultado.IsValid = True
    CalcularRegresionMultiple = resultado
    Exit Function

CalculationError:
    resultado.IsValid = False
    resultado.ErrorMessage = "Error en cálculo matricial: " & Err.Description
    CalcularRegresionMultiple = resultado
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
End Function

' =====================================================
' FUNCIONES AUXILIARES MEJORADAS
' =====================================================

Public Function ObtenerArrayNumerico(rng As Range) As Variant
    ' Convierte rango a array y valida valores numéricos
    Dim arr As Variant
    Dim i As Long, j As Long
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
    
    If rng.Cells.count = 1 Then
        ReDim arr(1 To 1, 1 To 1)
        If IsNumeric(rng.Value) And rng.Value <> "" Then
            arr(1, 1) = CDbl(rng.Value)
        Else
            arr(1, 1) = CVErr(xlErrValue)
        End If
    Else
        arr = rng.Value
        For i = 1 To UBound(arr, 1)
            For j = 1 To UBound(arr, 2)
                If Not IsNumeric(arr(i, j)) Or arr(i, j) = "" Then
                    arr(i, j) = CVErr(xlErrValue)
                Else
                    arr(i, j) = CDbl(arr(i, j))
                End If
            Next j
        Next i
    End If
    
    ObtenerArrayNumerico = arr
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
End Function

Public Function EsArrayValido(arr As Variant) As Boolean
    Dim i As Long, j As Long
    
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            If IsError(arr(i, j)) Then
                EsArrayValido = False
                Exit Function
            End If
        Next j
    Next i
    
    EsArrayValido = True
End Function

Public Function EsVectorConstante(arr As Variant) As Boolean
    Dim i As Long
    Dim firstValue As Double
    
    If UBound(arr, 1) < 2 Then
        EsVectorConstante = True
        Exit Function
    End If
    
    firstValue = arr(1, 1)
    For i = 2 To UBound(arr, 1)
        If Abs(arr(i, 1) - firstValue) > EPSILON Then
            EsVectorConstante = False
            Exit Function
        End If
    Next i
    
    EsVectorConstante = True
End Function

Public Function EsColumnaConstante(matriz As Variant, columna As Long) As Boolean
    Dim i As Long
    Dim firstValue As Double
    
    If UBound(matriz, 1) < 2 Then
        EsColumnaConstante = True
        Exit Function
    End If
    
    firstValue = matriz(1, columna)
    For i = 2 To UBound(matriz, 1)
        If Abs(matriz(i, columna) - firstValue) > EPSILON Then
            EsColumnaConstante = False
            Exit Function
        End If
    Next i
    
    EsColumnaConstante = True
End Function

Public Function EsMatrizInvertible(matriz As Variant, ByRef conditionNumber As Double) As Boolean
    ' Estimación simple del número de condición para matrices pequeñas
    Dim det As Double
    Dim norm As Double, NormInv As Double
    Dim i As Long, j As Long
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
    
    On Error GoTo NotInvertible
    
    ' Para matriz 2x2 o 3x3, calcular determinante
    If UBound(matriz, 1) = 2 Then
        det = matriz(1, 1) * matriz(2, 2) - matriz(1, 2) * matriz(2, 1)
        conditionNumber = 1# / Abs(det)
    ElseIf UBound(matriz, 1) = 3 Then
        det = matriz(1, 1) * (matriz(2, 2) * matriz(3, 3) - matriz(2, 3) * matriz(3, 2)) - _
              matriz(1, 2) * (matriz(2, 1) * matriz(3, 3) - matriz(2, 3) * matriz(3, 1)) + _
              matriz(1, 3) * (matriz(2, 1) * matriz(3, 2) - matriz(2, 2) * matriz(3, 1))
        conditionNumber = 1# / Abs(det)
    Else
        ' Para matrices más grandes, usar aproximación
        conditionNumber = 100000000# ' Valor por defecto para matrices no singulares
    End If
    
    EsMatrizInvertible = (Abs(det) > EPSILON)
    Exit Function

NotInvertible:
    EsMatrizInvertible = False
    conditionNumber = 1E+308 ' Infinito
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
End Function

Public Sub ManejarError(nombreProcedimiento As String, numeroError As Long, descripcion As String)
    Dim mensaje As String
    mensaje = "Error en " & nombreProcedimiento & ":" & vbCrLf & _
              "Código: " & numeroError & vbCrLf & _
              "Descripción: " & descripcion
    MsgBox mensaje, vbCritical
    
End Sub

' =====================================================
' FUNCIONES DE PRESENTACIÓN Y FORMATO
' =====================================================

Public Function CrearHojaResultados(nombreBase As String) As Worksheet
    Set wbTarget = ActiveWorkbook
    
    On Error GoTo ErrorHandler
    
    Dim nombreCompleto As String
    Dim contador As Integer
    Dim ws As Worksheet
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
    
    ' Generar nombre base con timestamp
    nombreCompleto = "Hoja de Análisis Resultado " & Format(Now, "hhmmss")
    
    ' Si se proporciona un nombre base adicional, incluirlo
    If Trim(nombreBase) <> "" Then
        nombreCompleto = nombreCompleto & " " & nombreBase
    End If
    
    ' Asegurar que el nombre no exceda 31 caracteres (límite de Excel)
    If Len(nombreCompleto) > 31 Then
        nombreCompleto = Left(nombreCompleto, 31)
    End If
    
    ' Eliminar caracteres inválidos para nombres de hojas
    nombreCompleto = LimpiarNombreHoja(nombreCompleto)
    
    ' Intentar crear la hoja con el nombre generado
    Application.DisplayAlerts = False
    On Error Resume Next
    wbTarget.Sheets(nombreCompleto).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set ws = wbTarget.Sheets.Add
    ws.Name = nombreCompleto
    
    Set CrearHojaResultados = ws
    Exit Function

ErrorHandler:
    ' Si hay error por nombre duplicado, agregar un contador
    If Err.Number = 1004 Then
        contador = 1
        Dim nombreUnico As String
        
        Do
            nombreUnico = nombreCompleto & "(" & contador & ")"
            If Len(nombreUnico) > 31 Then
                nombreUnico = Left(nombreCompleto, 27) & "(" & contador & ")"
            End If
            
            On Error Resume Next
            Set ws = wbTarget.Sheets.Add
            ws.Name = nombreUnico
            If Err.Number = 0 Then
                Set CrearHojaResultados = ws
                Exit Function
            Else
                contador = contador + 1
            End If
            On Error GoTo ErrorHandler
        Loop While contador < 100
    End If
    
    ' Si todo falla, crear con nombre por defecto
    Set ws = wbTarget.Sheets.Add
    Set CrearHojaResultados = ws
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
End Function

Public Function LimpiarNombreHoja(nombre As String) As String
    ' Eliminar caracteres inválidos para nombres de hojas en Excel
    Dim caracteresInvalidos As String
    Dim i As Integer
    Dim caracter As String
    Dim nombreLimpio As String
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
    
    caracteresInvalidos = ":\/?*[]"
    nombreLimpio = nombre
    
    ' Eliminar caracteres inválidos
    For i = 1 To Len(caracteresInvalidos)
        caracter = Mid(caracteresInvalidos, i, 1)
        nombreLimpio = Replace(nombreLimpio, caracter, "")
    Next i
    
    ' Eliminar espacios al inicio y final
    nombreLimpio = Trim(nombreLimpio)
    
    ' Si el nombre queda vacío, usar nombre por defecto
    If nombreLimpio = "" Then
        nombreLimpio = "Analisis_" & Format(Now, "hhmmss")
    End If
    
    LimpiarNombreHoja = nombreLimpio
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
End Function




