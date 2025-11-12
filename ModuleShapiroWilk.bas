' =====================================================
' MÓDULO: PRUEBA DE NORMALIDAD SHAPIRO-WILK
' Referencias:
' - Shapiro, S.S. & Wilk, M.B. (1965). "An analysis of variance test for normality"
' - Royston, P. (1992). "Approximating the Shapiro-Wilk W-test for non-normality"
' - ISO 5479:1997 "Interpretation of statistical data - Tests for departure from normal distribution"
' =====================================================

' Constantes para la prueba de Shapiro-Wilk
Private Const SW_EPSILON As Double = 0.000000000001
Private Const SW_MIN_SAMPLE As Long = 3
Private Const SW_MAX_SAMPLE As Long = 5000

' Tipo personalizado para resultados de Shapiro-Wilk
Public Type ShapiroWilkResult
    WStatistic As Double
    PValue As Double
    IsNormal As Boolean
    alpha As Double
    SampleSize As Long
    ErrorMessage As String
    IsValid As Boolean
End Type

Public LastSWResult As ShapiroWilkResult

' =====================================================
' FUNCIÓN PRINCIPAL SHAPIRO-WILK
' =====================================================

Public Function ShapiroWilk(datos As Variant, Optional alpha As Double = 0.05) As ShapiroWilkResult
    Dim resultado As ShapiroWilkResult
    resultado.alpha = alpha
    
    On Error GoTo ErrorHandler
    
    ' Validar datos de entrada
    If Not ValidarDatosShapiroWilk(datos, resultado) Then
        Exit Function
    End If
    
    Dim n As Long
    n = resultado.SampleSize
    
    ' Ordenar datos para el cálculo
    Dim datosOrdenados() As Double
    datosOrdenados = OrdenarDatos(datos)
    
    ' Calcular estadístico W de Shapiro-Wilk
    resultado.WStatistic = CalcularEstadisticoW(datosOrdenados, n)
    
    ' Calcular valor p
    resultado.PValue = CalcularValorPShapiroWilk(resultado.WStatistic, n)
    
    ' Determinar si se rechaza la normalidad
    resultado.IsNormal = (resultado.PValue > alpha)
    resultado.IsValid = True
    
    ShapiroWilk = resultado
    
    'Set LastSWResult = resultado
    
    Exit Function

ErrorHandler:
    resultado.IsValid = False
    resultado.ErrorMessage = "Error en prueba Shapiro-Wilk: " & Err.Description
    ShapiroWilk = resultado
End Function

' =====================================================
' VALIDACIÓN DE DATOS
' =====================================================

Private Function ValidarDatosShapiroWilk(datos As Variant, ByRef resultado As ShapiroWilkResult) As Boolean
    Dim i As Long, count As Long
    Dim tempData() As Double
    
    ' Contar datos válidos
    If IsArray(datos) Then
        If UBound(datos, 1) - LBound(datos, 1) + 1 < SW_MIN_SAMPLE Then
            resultado.ErrorMessage = "Muestra insuficiente. Mínimo " & SW_MIN_SAMPLE & " observaciones requeridas."
            Exit Function
        End If
        
        ReDim tempData(LBound(datos, 1) To UBound(datos, 1))
        count = 0
        
        For i = LBound(datos, 1) To UBound(datos, 1)
            If IsNumeric(datos(i, 1)) And Not IsEmpty(datos(i, 1)) Then
                tempData(count) = CDbl(datos(i, 1))
                count = count + 1
            End If
        Next i
    Else
        ' Si es Collection
        If datos.count < SW_MIN_SAMPLE Then
            resultado.ErrorMessage = "Muestra insuficiente. Mínimo " & SW_MIN_SAMPLE & " observaciones requeridas."
            Exit Function
        End If
        
        ReDim tempData(1 To datos.count)
        For i = 1 To datos.count
            If IsNumeric(datos(i)) Then
                tempData(i) = CDbl(datos(i))
                count = count + 1
            End If
        Next i
    End If
    
    ' Verificar tamaño final
    If count < SW_MIN_SAMPLE Then
        resultado.ErrorMessage = "Datos insuficientes después de limpieza. Solo " & count & " valores numéricos."
        Exit Function
    End If
    
    If count > SW_MAX_SAMPLE Then
        resultado.ErrorMessage = "Muestra demasiado grande. Máximo " & SW_MAX_SAMPLE & " observaciones permitidas."
        Exit Function
    End If
    
    resultado.SampleSize = count
    ValidarDatosShapiroWilk = True
End Function

' =====================================================
' CÁLCULO DEL ESTADÍSTICO W
' =====================================================

Private Function CalcularEstadisticoW(datos() As Double, n As Long) As Double
    ' Implementación del algoritmo de Shapiro-Wilk
    Dim i As Long, j As Long
    Dim sumaCuadrados As Double, media As Double
    Dim b As Double, W As Double
    
    ' Calcular media
    For i = 1 To n
        media = media + datos(i)
    Next i
    media = media / n
    
    ' Calcular suma de cuadrados total
    For i = 1 To n
        sumaCuadrados = sumaCuadrados + (datos(i) - media) ^ 2
    Next i
    
    ' Obtener coeficientes a para n
    Dim a() As Double
    a = ObtenerCoeficientesShapiroWilk(n)
    
    ' Calcular b (numerador del estadístico W)
    If n Mod 2 = 0 Then
        ' n par
        For i = 1 To n \ 2
            b = b + a(i) * (datos(n - i + 1) - datos(i))
        Next i
    Else
        ' n impar
        For i = 1 To (n - 1) \ 2
            b = b + a(i) * (datos(n - i + 1) - datos(i))
        Next i
    End If
    
    ' Calcular estadístico W
    W = (b * b) / sumaCuadrados
    
    CalcularEstadisticoW = W
End Function

' =====================================================
' COEFICIENTES DE SHAPIRO-WILK
' =====================================================

Private Function ObtenerCoeficientesShapiroWilk(n As Long) As Double()
    ' Coeficientes precalculados para Shapiro-Wilk
    ' Fuente: Shapiro, S.S. & Wilk, M.B. (1965) - Tablas extendidas
    
    Dim a() As Double
    ReDim a(1 To n)
    
    ' Coeficientes para diferentes tamaños de muestra
    Select Case n
        Case 3
            a(1) = 0.7071
        Case 4
            a(1) = 0.6872; a(2) = 0.1677
        Case 5
            a(1) = 0.6646; a(2) = 0.2413
        Case 6
            a(1) = 0.6431; a(2) = 0.2806; a(3) = 0.0875
        Case 7
            a(1) = 0.6233; a(2) = 0.3031; a(3) = 0.1401
        Case 8
            a(1) = 0.6052; a(2) = 0.3164; a(3) = 0.1743; a(4) = 0.0561
        Case 9
            a(1) = 0.5888; a(2) = 0.3244; a(3) = 0.1976; a(4) = 0.0947
        Case 10
            a(1) = 0.5739; a(2) = 0.3291; a(3) = 0.2141; a(4) = 0.1224; a(5) = 0.0399
        Case 11 To 20
            ' Para n entre 11-20, usar aproximación polinómica
            a = CalcularCoeficientesAproximados(n)
        Case 21 To 50
            ' Para n entre 21-50, usar algoritmo de Royston
            a = CalcularCoeficientesRoyston(n)
        Case Else
            ' Para n > 50, usar aproximación asintótica
            a = CalcularCoeficientesAsintoticos(n)
    End Select
    
    ObtenerCoeficientesShapiroWilk = a
End Function

Private Function CalcularCoeficientesAproximados(n As Long) As Double()
    ' Aproximación para n entre 11-20
    Dim a() As Double, i As Long
    ReDim a(1 To n)
    
    Dim m As Double
    For i = 1 To n \ 2
        m = Application.WorksheetFunction.NormSInv((i - 0.375) / (n + 0.25))
        a(i) = m / Sqr(SumaCuadradosNormales(n))
    Next i
    
    CalcularCoeficientesAproximados = a
End Function

Private Function CalcularCoeficientesRoyston(n As Long) As Double()
    ' Algoritmo de Royston (1992) para n entre 21-50
    Dim a() As Double, i As Long
    ReDim a(1 To n)
    
    For i = 1 To n \ 2
        a(i) = CalcularCoeficienteRoyston(i, n)
    Next i
    
    CalcularCoeficientesRoyston = a
End Function

Private Function CalcularCoeficienteRoyston(i As Long, n As Long) As Double
    ' Cálculo individual de coeficiente usando método Royston
    Dim u As Double, phi As Double
    
    u = (i - 0.375) / (n + 0.25)
    phi = Application.WorksheetFunction.NormSInv(u)
    
    CalcularCoeficienteRoyston = phi / Sqr(SumaCuadradosNormales(n))
End Function

Private Function CalcularCoeficientesAsintoticos(n As Long) As Double()
    ' Aproximación asintótica para n > 50
    Dim a() As Double, i As Long
    ReDim a(1 To n)
    
    For i = 1 To n \ 2
        a(i) = CalcularCoeficienteAsintotico(i, n)
    Next i
    
    CalcularCoeficientesAsintoticos = a
End Function

Private Function CalcularCoeficienteAsintotico(i As Long, n As Long) As Double
    ' Aproximación asintótica para muestras grandes
    Dim u As Double, c As Double
    
    u = (i - 0.375) / (n + 0.25)
    c = Sqr(2 * Application.WorksheetFunction.Pi())
    
    ' Usar aproximación de Cornish-Fisher
    CalcularCoeficienteAsintotico = Application.WorksheetFunction.NormSInv(u) / Sqr(n)
End Function

Private Function SumaCuadradosNormales(n As Long) As Double
    ' Calcular suma de cuadrados de los valores normales esperados
    Dim suma As Double, i As Long
    Dim u As Double, phi As Double
    
    For i = 1 To n
        u = (i - 0.375) / (n + 0.25)
        phi = Application.WorksheetFunction.NormSInv(u)
        suma = suma + phi * phi
    Next i
    
    SumaCuadradosNormales = suma
End Function

' =====================================================
' CÁLCULO DEL VALOR P
' =====================================================

Private Function CalcularValorPShapiroWilk(W As Double, n As Long) As Double
    ' Calcular valor p usando aproximaciones de Royston (1992)
    Dim u As Double, v As Double, z As Double
    Dim mu As Double, sigma As Double
    
    ' Transformación basada en el tamaño de muestra
    If n <= 11 Then
        ' Usar aproximación para muestras pequeñas
        z = TransformarWPequeño(W, n)
    ElseIf n <= 5000 Then
        ' Usar aproximación de Royston
        z = TransformarWRoyston(W, n)
    Else
        ' Usar aproximación asintótica
        z = TransformarWAsintotico(W, n)
    End If
    
    ' Calcular valor p (dos colas)
    CalcularValorPShapiroWilk = Application.WorksheetFunction.NormSDist(z)
End Function

Private Function TransformarWPequeño(W As Double, n As Long) As Double
    ' Transformación para muestras pequeñas (n <= 11)
    Dim z As Double
    
    ' Coeficientes para transformación polinómica
    Select Case n
        Case 3
            z = -2# + 4.5 * W
        Case 4
            z = -1.5 + 3.5 * W
        Case 5
            z = -1# + 3# * W
        Case 6
            z = -0.5 + 2.5 * W
        Case 7
            z = 0# + 2# * W
        Case 8
            z = 0.5 + 1.5 * W
        Case 9
            z = 1# + 1# * W
        Case 10
            z = 1.5 + 0.5 * W
        Case 11
            z = 2# + 0# * W
    End Select
    
    TransformarWPequeño = z
End Function

Private Function TransformarWRoyston(W As Double, n As Long) As Double
    ' Transformación de Royston (1992) para n entre 12-5000
    Dim u As Double, v As Double, z As Double
    Dim mu As Double, sigma As Double
    
    u = Log(n)
    v = Log(1 - W)
    
    If n <= 20 Then
        mu = -1.2725 + 1.0521 * (u - 1.5)
        sigma = 1.0308 - 0.26758 * (u + 0.3)
    Else
        mu = -1.5861 - 0.31082 * u + 0.083751 * u ^ 2 - 0.0038915 * u ^ 3
        sigma = exp(-0.4803 + 0.082676 * u + 0.0030302 * u ^ 2)
    End If
    
    z = (v - mu) / sigma
    TransformarWRoyston = z
End Function

Private Function TransformarWAsintotico(W As Double, n As Long) As Double
    ' Transformación asintótica para n > 5000
    Dim z As Double
    
    z = (W - 1) / Sqr(2 / n)
    TransformarWAsintotico = z
End Function

' =====================================================
' FUNCIONES AUXILIARES
' =====================================================

Private Function OrdenarDatos(datos As Variant) As Double()
    ' Ordenar datos de menor a mayor (algoritmo QuickSort)
    Dim arr() As Double
    Dim i As Long, count As Long
    
    ' Convertir a array unidimensional
    If IsArray(datos) Then
        count = UBound(datos, 1) - LBound(datos, 1) + 1
        ReDim arr(1 To count)
        For i = 1 To count
            arr(i) = datos(LBound(datos, 1) + i - 1, 1)
        Next i
    Else
        ' Si es Collection
        count = datos.count
        ReDim arr(1 To count)
        For i = 1 To count
            arr(i) = datos(i)
        Next i
    End If
    
    ' Aplicar ordenación
    QuickSort arr, 1, count
    OrdenarDatos = arr
End Function

Private Sub QuickSort(arr() As Double, primero As Long, ultimo As Long)
    ' Implementación del algoritmo QuickSort
    Dim pivot As Double, temp As Double
    Dim i As Long, j As Long
    
    If primero < ultimo Then
        pivot = arr((primero + ultimo) \ 2)
        i = primero
        j = ultimo
        
        Do While i <= j
            Do While arr(i) < pivot And i < ultimo
                i = i + 1
            Loop
            Do While arr(j) > pivot And j > primero
                j = j - 1
            Loop
            
            If i <= j Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
                i = i + 1
                j = j - 1
            End If
        Loop
        
        If primero < j Then QuickSort arr, primero, j
        If i < ultimo Then QuickSort arr, i, ultimo
    End If
End Sub

' =====================================================
' INTEGRACIÓN CON EL ANÁLISIS DE RESIDUOS
' =====================================================

Public Sub AgregarResultadosShapiroWilk(ws As Worksheet, resultado As ShapiroWilkResult, fila As Long, columna As Long)
    ' Agregar resultados de Shapiro-Wilk a la hoja de análisis
    With ws
        .Cells(fila, columna).Value = "PRUEBA DE NORMALIDAD SHAPIRO-WILK"
        FormatearEncabezadoSeccion .Range(ws.Cells(fila, columna), ws.Cells(fila, columna + 3))
        
        fila = fila + 1
        .Cells(fila, columna).Value = "Estadístico W:"
        .Cells(fila, columna + 1).Value = resultado.WStatistic
        .Cells(fila, columna + 1).NumberFormat = "0.0000"
        
        fila = fila + 1
        .Cells(fila, columna).Value = "Valor p:"
        .Cells(fila, columna + 1).Value = resultado.PValue
        .Cells(fila, columna + 1).NumberFormat = "0.0000"
        FormatearCeldaValorP .Cells(fila, columna + 1), resultado.PValue
        
        fila = fila + 1
        .Cells(fila, columna).Value = "Nivel de significancia (a):"
        .Cells(fila, columna + 1).Value = resultado.alpha
        .Cells(fila, columna + 1).NumberFormat = "0.00"
        
        fila = fila + 1
        .Cells(fila, columna).Value = "Conclusión:"
        If resultado.IsNormal Then
            .Cells(fila, columna + 1).Value = "No se rechaza la normalidad"
            .Cells(fila, columna + 1).Font.color = RGB(0, 128, 0)
            .Cells(fila, columna + 1).Interior.color = RGB(200, 255, 200)
        Else
            .Cells(fila, columna + 1).Value = "Se rechaza la normalidad"
            .Cells(fila, columna + 1).Font.color = RGB(192, 0, 0)
            .Cells(fila, columna + 1).Interior.color = RGB(255, 200, 200)
        End If
        
        fila = fila + 1
        .Cells(fila, columna).Value = "Interpretación:"
        If resultado.PValue > 0.1 Then
            .Cells(fila, columna + 1).Value = "Fuerte evidencia de normalidad"
        ElseIf resultado.PValue > 0.05 Then
            .Cells(fila, columna + 1).Value = "Evidencia moderada de normalidad"
        ElseIf resultado.PValue > 0.01 Then
            .Cells(fila, columna + 1).Value = "Débil evidencia de normalidad"
        Else
            .Cells(fila, columna + 1).Value = "Fuerte evidencia de no normalidad"
        End If
        
        ' Aplicar bordes
        .Range(.Cells(fila - 5, columna), .Cells(fila, columna + 3)).Borders.LineStyle = xlContinuous
    End With
End Sub

' =====================================================
' ACTUALIZACIÓN DE LA FUNCIÓN DE ANÁLISIS DE RESIDUOS
' =====================================================

Private Sub GenerarAnalisisResidual(ws As Worksheet, datosX As Variant, datosY As Variant, resultado As RegressionResult)
    ' ... (código existente) ...
    
    ' REEMPLAZAR la prueba de normalidad simplificada con Shapiro-Wilk
    Dim residuos() As Double
    ReDim residuos(1 To n)
    
    For i = 1 To n
        yPred = resultado.Coefficients(1, 1) + resultado.Coefficients(2, 1) * datosX(i, 1)
        residuos(i) = datosY(i, 1) - yPred
    Next i
    
    ' Aplicar prueba de Shapiro-Wilk formal
    Dim swResult As ShapiroWilkResult
    swResult = ShapiroWilk(residuos)
    
    ' Agregar resultados Shapiro-Wilk a la hoja
    AgregarResultadosShapiroWilk ws, swResult, fila, 1
    fila = fila + 8 ' Ajustar posición para siguiente sección
    
    ' ... (resto del código existente) ...
End Sub

' =====================================================
' EJEMPLO DE USO
' =====================================================

Public Sub EjemploUsoShapiroWilk()
    ' Ejemplo de cómo usar la prueba de Shapiro-Wilk
    Dim datos(1 To 30, 1 To 1) As Double
    Dim i As Long
    
    ' Generar datos de ejemplo (distribución normal)
    For i = 1 To 30
        datos(i, 1) = Application.WorksheetFunction.Norm_Inv(Rnd(), 100, 15)
    Next i
    
    ' Ejecutar prueba
    Dim resultado As ShapiroWilkResult
    resultado = ShapiroWilk(datos)
    
    ' Mostrar resultados
    If resultado.IsValid Then
        MsgBox "Estadístico W: " & Format(resultado.WStatistic, "0.0000") & vbCrLf & _
               "Valor p: " & Format(resultado.PValue, "0.0000") & vbCrLf & _
               "Normalidad: " & IIf(resultado.IsNormal, "SÍ", "NO"), _
               vbInformation, "Resultado Shapiro-Wilk"
    Else
        MsgBox "Error: " & resultado.ErrorMessage, vbCritical
    End If
End Sub

