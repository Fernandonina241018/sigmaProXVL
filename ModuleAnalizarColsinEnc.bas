
'================================================================================
' FUNCIÓN: AnalizarColumnaSinEncabezado
' PROPÓSITO: Alternativa para análisis cuando primera fila NO es encabezado
' DIFERENCIAS: Incluye primera fila en análisis, nombre genérico para columna
' USO: Para datos que no contienen fila de encabezado
'================================================================================
Function AnalizarColumnaSinEncabezado(rangoColumna As Range) As EstadisticasColumna
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
        .ScreenUpdating = False             ' No actualizar pantalla (CRÍTICO)
        .Calculation = xlCalculationManual  ' Desactivar cálculos automáticos
        .EnableEvents = False               ' Desactivar eventos
        .DisplayStatusBar = False           ' Ocultar barra de estado
    End With
    
    On Error GoTo Cleanup
    
    ' Inicialización idéntica a función principal
    resultados.count = 0
    suma = 0
    sumaCuadrados = 0
    resultados.maximo = -1.79769313486231E+308
    resultados.minimo = 1.79769313486231E+308
    resultados.rango = rangoColumna.Address
    
    ' Metadata - nombre genérico
    On Error Resume Next
    resultados.columna = rangoColumna.Cells(1, 1).EntireColumn.Address(0, 0)
    resultados.columna = Replace(resultados.columna, ":", "")
    resultados.NombreColumna = "Columna " & resultados.columna ' Nombre por defecto
    On Error GoTo 0
    
    ' DIFERENCIA PRINCIPAL: Incluir primera fila en análisis
    For Each celda In rangoColumna.Cells
        If IsNumeric(celda.Value) And celda.Value <> "" Then
            ReDim Preserve valores(resultados.count)
            valores(resultados.count) = celda.Value
            suma = suma + celda.Value
            resultados.count = resultados.count + 1
        End If
    Next celda
    
    ' Cálculos estadísticos (idénticos a función principal)
    If resultados.count > 0 Then
        resultados.promedio = suma / resultados.count
        
        For i = 0 To resultados.count - 1
            sumaCuadrados = sumaCuadrados + (valores(i) - resultados.promedio) ^ 2
        Next i
        
        If resultados.count > 1 Then
            resultados.desviacionEstandar = Sqr(sumaCuadrados / (resultados.count - 1))
        Else
            resultados.desviacionEstandar = 0
        End If
        
        If resultados.promedio <> 0 Then
            resultados.RSD = CalcularRSD(valores, True)
        Elsevalores
            resultados.RSD = 0
        End If
        
        If resultados.promedio <> 0 Then
            resultados.varianza = CalcularVarianza()
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
        
        For i = 0 To resultados.count - 1
            If valores(i) > resultados.maximo Then resultados.maximo = valores(i)
            If valores(i) < resultados.minimo Then resultados.minimo = valores(i)
        Next i
        
        ' Incluir detección de outliers
        CalcularOutliersIQR valores, resultados
    Else
        resultados.NumOutliers = 0
        resultados.MediaRobusta = 0
        resultados.DesvEstandarRobusta = 0
        resultados.RSDrobusto = 0
    End If
    
    AnalizarColumnaSinEncabezado = resultados
    
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
    
End Function

Public Function CalcularMediana(valores() As Double) As Double
    Dim temp() As Double
    Dim n As Long, mitad As Long
    Dim i As Long, j As Long, tmp As Double
    
    n = UBound(valores) - LBound(valores) + 1
    ReDim temp(0 To n - 1)
    
    ' Copiar valores
    For i = 0 To n - 1
        temp(i) = valores(i)
    Next i
    
    ' Ordenar (burbuja simple)
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If temp(i) > temp(j) Then
                tmp = temp(i)
                temp(i) = temp(j)
                temp(j) = tmp
            End If
        Next j
    Next i
    
    mitad = n \ 2
    If n Mod 2 = 0 Then
        CalcularMediana = (temp(mitad - 1) + temp(mitad)) / 2
    Else
        CalcularMediana = temp(mitad)
    End If
    
End Function

Public Function CalcularVarianza(valores() As Double) As Double
    Dim promedio As Double
    Dim suma As Double
    Dim sumaCuadrados As Double
    Dim i As Long
    Dim n As Long

    n = UBound(valores) - LBound(valores) + 1
    If n <= 1 Then
        CalcularVarianza = 0
        Exit Function
    End If

    ' Calcular promedio
    For i = 0 To n - 1
        suma = suma + valores(i)
    Next i
    promedio = suma / n

    ' Calcular suma de cuadrados
    For i = 0 To n - 1
        sumaCuadrados = sumaCuadrados + (valores(i) - promedio) ^ 2
    Next i

    ' Varianza muestral
    CalcularVarianza = sumaCuadrados / (n - 1)
End Function

Function CalcularModa(valores() As Double) As Double
    Dim dict As Object
    Dim i As Long
    Dim valor As Variant
    Dim maxFrecuencia As Long
    Dim moda As Double

    Set dict = CreateObject("Scripting.Dictionary")

    ' Contar frecuencia de cada valor
    For i = LBound(valores) To UBound(valores)
        If dict.Exists(valores(i)) Then
            dict(valores(i)) = dict(valores(i)) + 1
        Else
            dict.Add valores(i), 1
        End If
    Next i

    ' Buscar el valor con mayor frecuencia
    maxFrecuencia = 0
    For Each valor In dict.Keys
        If dict(valor) > maxFrecuencia Then
            maxFrecuencia = dict(valor)
            moda = valor
        End If
    Next valor

    CalcularModa = moda
End Function
