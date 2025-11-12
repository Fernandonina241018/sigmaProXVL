'================================================================================
' FUNCIÓN: CalcularAsimetria
' PROPÓSITO: Calcula coeficiente de asimetría de Fisher (g1) - Compatible Excel
' FÓRMULA: g1 = [n/((n-1)(n-2))] × S[(x?-x¯)/s]³
' RETORNA: Double (-8 a +8)
'          Negativo = Asimetría izquierda
'          Cero     = Simétrica
'          Positivo = Asimetría derecha
'================================================================================
Public Function CalcularAsimetria(valores() As Double) As Double
    On Error GoTo ErrorHandler
    
    Dim n As Long
    Dim i As Long
    Dim promedio As Double
    Dim suma As Double
    Dim sumaCuadrados As Double
    Dim sumaCubos As Double
    Dim desviacionEstandar As Double
    Dim valorEstandarizado As Double
    Dim sumaZ3 As Double
    
    '----------------------------------------------------------------------------
    ' VALIDACIÓN 1: Tamaño de muestra
    '----------------------------------------------------------------------------
    n = UBound(valores) - LBound(valores) + 1
    
    ' Mínimo 3 observaciones para asimetría de Fisher
    If n < 3 Then
        CalcularAsimetria = 0  ' ?? Considera lanzar error en producción
        Exit Function
    End If
    
    '----------------------------------------------------------------------------
    ' PASO 1: Calcular promedio
    '----------------------------------------------------------------------------
    suma = 0
    For i = LBound(valores) To UBound(valores)
        ' Validar valores individuales
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
    ' Usar epsilon para evitar problemas de precisión
    Const EPSILON As Double = 0.0000001
    
    If Abs(desviacionEstandar) < EPSILON Then
        ' Todos los valores son iguales ? Asimetría = 0
        CalcularAsimetria = 0
        Exit Function
    End If
    
    '----------------------------------------------------------------------------
    ' PASO 3: Calcular suma de valores estandarizados al cubo
    '----------------------------------------------------------------------------
    sumaZ3 = 0
    For i = LBound(valores) To UBound(valores)
        valorEstandarizado = (valores(i) - promedio) / desviacionEstandar
        sumaZ3 = sumaZ3 + valorEstandarizado ^ 3
    Next i
    
    '----------------------------------------------------------------------------
    ' PASO 4: Aplicar factor de corrección de Fisher
    '----------------------------------------------------------------------------
    ' ? FÓRMULA CORRECTA: Compatible con Excel SKEW()
    CalcularAsimetria = (n / ((n - 1) * (n - 2))) * sumaZ3
    
    Exit Function

ErrorHandler:
    ' En caso de error, retornar 0 (o considera lanzar error)
    CalcularAsimetria = 0
    Debug.Print "Error en CalcularAsimetria: " & Err.Description
End Function
