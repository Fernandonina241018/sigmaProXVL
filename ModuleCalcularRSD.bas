'   - También conocido como Coeficiente de Variación (CV)
'   - Usado en control de calidad farmacéutico y químico
'   - FDA recomienda RSD < 2% para métodos analíticos
'   - ICH Q2(R1) acepta RSD < 15% para ensayos de contenido
'================================================================================
Public Function CalcularRSD(valores() As Double, Optional usarMuestral As Boolean = True) As Double
    On Error GoTo ErrorHandler
    
    Dim n As Long
    Dim promedio As Double
    Dim desviacionEstandar As Double
    Dim i As Long
    Dim suma As Double
    Dim sumaCuadrados As Double
    Dim divisor As Long
    
    '----------------------------------------------------------------------------
    ' VALIDACIÓN 1: Tamaño de muestra
    '----------------------------------------------------------------------------
    n = UBound(valores) - LBound(valores) + 1
    
    If n < 2 Then
        CalcularRSD = 0
        Exit Function
    End If
    
    '----------------------------------------------------------------------------
    ' PASO 1: Calcular promedio (media aritmética)
    '----------------------------------------------------------------------------
    suma = 0
    For i = LBound(valores) To UBound(valores)
        ' Validar que sea numérico
        If Not IsNumeric(valores(i)) Then GoTo ErrorHandler
        suma = suma + valores(i)
    Next i
    
    promedio = suma / n
    
    '----------------------------------------------------------------------------
    ' VALIDACIÓN 2: Promedio no puede ser cero
    '----------------------------------------------------------------------------
    Const EPSILON As Double = 0.0000001
    
    If Abs(promedio) < EPSILON Then
        ' RSD indefinido cuando la media es cero
        CalcularRSD = 0
        Exit Function
    End If
    
    '----------------------------------------------------------------------------
    ' PASO 2: Calcular desviación estándar
    '----------------------------------------------------------------------------
    sumaCuadrados = 0
    For i = LBound(valores) To UBound(valores)
        sumaCuadrados = sumaCuadrados + (valores(i) - promedio) ^ 2
    Next i
    
    ' Determinar divisor según tipo de desviación
    If usarMuestral Then
        divisor = n - 1  ' Desviación estándar MUESTRAL (s)
    Else
        divisor = n      ' Desviación estándar POBLACIONAL (s)
    End If
    
    ' Evitar división por cero
    If divisor = 0 Then
        CalcularRSD = 0
        Exit Function
    End If
    
    desviacionEstandar = Sqr(sumaCuadrados / divisor)
    
    '----------------------------------------------------------------------------
    ' PASO 3: Calcular RSD como porcentaje
    '----------------------------------------------------------------------------
    ' RSD(%) = (s / x¯) × 100
    CalcularRSD = (desviacionEstandar / Abs(promedio)) * 100
    
    Exit Function

ErrorHandler:
    CalcularRSD = 0
    Debug.Print "Error en CalcularRSD: " & Err.Description
End Function

