'================================================================================
' FUNCIÓN: CalcularCurtosis
' PROPÓSITO: Calcula coeficiente de curtosis de Fisher (g2) - Compatible Excel
' FÓRMULA: g2 = [n(n+1)/((n-1)(n-2)(n-3))] × [S((x?-x¯)/s)4] - [3(n-1)²/((n-2)(n-3))]
' RETORNA: Double
'          < 0 = Platicúrtica (más plana que normal)
'          = 0 = Mesocúrtica (distribución normal)
'          > 0 = Leptocúrtica (más puntiaguda que normal)
' NOTA: Esta es la curtosis de EXCESO (ya ajustada con -3)
'================================================================================
Public Function CalcularCurtosis(valores() As Double) As Double
    On Error GoTo ErrorHandler
    
    Dim n As Long
    Dim i As Long
    Dim promedio As Double
    Dim suma As Double
    Dim sumaCuadrados As Double
    Dim desviacionEstandar As Double
    Dim valorEstandarizado As Double
    Dim sumaZ4 As Double
    Dim factor1 As Double
    Dim factor2 As Double
    
    '----------------------------------------------------------------------------
    ' VALIDACIÓN 1: Tamaño de muestra
    '----------------------------------------------------------------------------
    n = UBound(valores) - LBound(valores) + 1
    
    ' ? CORRECCIÓN: Curtosis de Fisher requiere mínimo 4 observaciones
    If n < 4 Then
        CalcularCurtosis = 0  ' O considera lanzar error
        Exit Function
    End If
    
    '----------------------------------------------------------------------------
    ' PASO 1: Calcular promedio
    '----------------------------------------------------------------------------
    suma = 0
    For i = LBound(valores) To UBound(valores)
        If Not IsNumeric(valores(i)) Then GoTo ErrorHandler
        suma = suma + valores(i)
    Next i
    promedio = suma / n
    
    '----------------------------------------------------------------------------
    ' PASO 2: Calcular desviación estándar MUESTRAL (n-1)
    '----------------------------------------------------------------------------
    sumaCuadrados = 0
    For i = LBound(valores) To UBound(valores)
        sumaCuadrados = sumaCuadrados + (valores(i) - promedio) ^ 2
    Next i
    
    ' ? CORRECCIÓN CRÍTICA: Dividir entre (n-1), no n
    desviacionEstandar = Sqr(sumaCuadrados / (n - 1))
    
    '----------------------------------------------------------------------------
    ' VALIDACIÓN 2: Desviación estándar válida
    '----------------------------------------------------------------------------
    Const EPSILON As Double = 0.0000001
    
    If Abs(desviacionEstandar) < EPSILON Then
        ' Todos los valores son iguales ? Curtosis indefinida
        CalcularCurtosis = 0
        Exit Function
    End If
    
    '----------------------------------------------------------------------------
    ' PASO 3: Calcular suma de valores estandarizados a la cuarta potencia
    '----------------------------------------------------------------------------
    sumaZ4 = 0
    For i = LBound(valores) To UBound(valores)
        valorEstandarizado = (valores(i) - promedio) / desviacionEstandar
        sumaZ4 = sumaZ4 + valorEstandarizado ^ 4
    Next i
    
    '----------------------------------------------------------------------------
    ' PASO 4: Aplicar fórmula de Fisher con factores de corrección
    '----------------------------------------------------------------------------
    ' ? FÓRMULA CORRECTA: Compatible con Excel KURT()
    
    ' Factor 1: [n(n+1)] / [(n-1)(n-2)(n-3)]
    factor1 = (n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))
    
    ' Factor 2: [3(n-1)²] / [(n-2)(n-3)]
    factor2 = (3 * (n - 1) ^ 2) / ((n - 2) * (n - 3))
    
    ' Curtosis de Fisher (exceso de curtosis)
    CalcularCurtosis = (factor1 * sumaZ4) - factor2
    
    Exit Function

ErrorHandler:
    CalcularCurtosis = 0
    Debug.Print "Error en CalcularCurtosis: " & Err.Description
End Function
