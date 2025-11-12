' Módulo: ProcCapabilityAnalysis
' Propósito: Análisis de Capacidad del Proceso para auditorías internacionales.
' Requisitos: Inputs desde UserForm (LSE, LIE, Target, Rango de datos via RefEdit).
' Referencias: ISO 22514-2, IATF 16949, AIAG SPC Manual.
' Autor: [Tu Nombre]
' Fecha: [Fecha]
' Historial de Revisiones: [Fecha, Cambios, Razón]

Option Explicit

' Constantes para pruebas de normalidad
Private Const alpha As Double = 0.05 ' Nivel de significancia para Shapiro-Wilk

Public Sub RunCapabilityAnalysis()
    On Error GoTo ErrorHandler
    Dim wsData As Worksheet, wsResults As Worksheet
    Dim dataRange As Range, targetCell As Range, adjustedRange As Range
    Dim LSE As Double, LIE As Double, Target As Double
    Dim dataPoints As Collection
    Dim mean As Double, stdDevOverall As Double, stdDevWithin As Double
    Dim Cp As Double, Cpk As Double, Pp As Double, Ppk As Double, Cpm As Double
    Dim IsNormal As Boolean
    
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
    
    ' 1. OBTENER INPUTS DEL USERFORM
    ' Asume: UserForm con txtLSE, txtLIE, txtTarget, refDataRange
    LSE = sigmaproxvl.cboLimiteSuperior.Value
    LIE = sigmaproxvl.cboLimiteInferior.Value
    Target = sigmaproxvl.cboExpectativa.Value
    Set dataRange = Range(sigmaproxvl.txtRango.Value)
    
    
    ' Excluir la primera fila del rango seleccionado
    Set adjustedRange = dataRange.Offset(1, 0).Resize(dataRange.Rows.count - 1, dataRange.Columns.count)

    ' Validar inputs críticos
    If LSE <= LIE Then
        Debug.Print "Error: LSE debe ser mayor que LIE.", vbCritical
        Exit Sub
    End If
    
    ' 2. LEER DATOS Y CALCULAR ESTADÍSTICAS
    Set dataPoints = ReadDataPoints(adjustedRange)
    mean = CalculateMean(dataPoints)
    stdDevOverall = CalculateStdDevOverall(dataPoints)
    ' Si hay subgrupos, calcular stdDevWithin desde R-bar/d2; si no, usar stdDevOverall
    stdDevWithin = EstimateStdDevWithin(adjustedRange) ' Implementar basado en subgrupos
    
    ' 3. VERIFICAR NORMALIDAD (Shapiro-Wilk)
    IsNormal = ShapiroWilkTest(dataPoints) ' Implementar o usar otra prueba
    
    ' 4. CALCULAR ÍNDICES DE CAPACIDAD
    If stdDevWithin > 0 Then
        Cp = (LSE - LIE) / (6 * stdDevWithin)
        Cpk = Application.WorksheetFunction.Min((LSE - mean) / (3 * stdDevWithin), (mean - LIE) / (3 * stdDevWithin))
    Else
        Cp = 0
        Cpk = 0
    End If
    
    If stdDevOverall > 0 Then
        Pp = (LSE - LIE) / (6 * stdDevOverall)
        Ppk = Application.WorksheetFunction.Min((LSE - mean) / (3 * stdDevOverall), (mean - LIE) / (3 * stdDevOverall))
    Else
        Pp = 0
        Ppk = 0
    End If

        
    If Target <> 0 And stdDevOverall > 0 Then
        Cpm = (LSE - LIE) / (6 * Sqr(stdDevOverall ^ 2 + (mean - Target) ^ 2))
    Else
        Cpm = 0
    End If

    
    ' 5. CREAR HOJA DE RESULTADOS
    Set wsResults = CreateUniqueSheet("Capacidad_Resultados")
    PopulateResultsSheet wsResults, dataPoints.count, mean, stdDevOverall, _
        stdDevWithin, Cp, Cpk, Pp, Ppk, Cpm, IsNormal, LSE, LIE, Target
    
    ' 6. GENERAR GRÁFICO
    GenerateHistogram wsResults, dataPoints, LSE, LIE, Target
    
    Debug.Print "Análisis de capacidad completado.", vbInformation
    Debug.Print "Media: " & mean
    Debug.Print "LSE: " & LSE & " | LIE: " & LIE
    Debug.Print "Cpk Numerador 1: " & (LSE - mean)
    Debug.Print "Cpk Numerador 2: " & (mean - LIE)
    Debug.Print "StdDevWithin: " & stdDevWithin

    Exit Sub
    
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With

ErrorHandler:
    Debug.Print "Error en análisis de capacidad: " & Now() & Err.Description, vbCritical
End Sub

' Función para leer datos desde rango
Private Function ReadDataPoints(dataRange As Range) As Collection
    Dim cell As Range, points As New Collection
    For Each cell In dataRange
        If IsNumeric(cell.Value) Then points.Add cell.Value
    Next cell
    Set ReadDataPoints = points
End Function

' Función para calcular media
Private Function CalculateMean(points As Collection) As Double
    Dim sum As Double, i As Long, count As Long
    If points.count <= 1 Then
        CalculateMean = 0
        Exit Function
    End If
    
    For i = 1 To points.count
        If IsNumeric(points(i)) Then
            sum = sum + CDbl(points(i))
            count = count + 1
        End If
    Next i
    
    If count = 0 Then
        CalculateMean = 0
    Else
        CalculateMean = sum / count
    End If
End Function

Private Function CalculateStdDevOverall(points As Collection) As Double
    Dim mean As Double, sumSq As Double
    Dim i As Long
    Dim n As Long
    Dim x As Double

    ' Validar que haya suficientes datos
    n = points.count
    If n < 1 Then
        CalculateStdDevOverall = 0
        Exit Function
    End If

    ' Calcular la media
    mean = 0
    For i = 1 To n
        If IsNumeric(points(i)) Then
            mean = mean + CDbl(points(i))
        Else
            Debug.Print "Dato no numérico en posición " & i & ": " & points(i)
        End If
    Next i
    mean = mean / n

    ' Calcular suma de cuadrados
    sumSq = 0
    For i = 1 To n
        If IsNumeric(points(i)) Then
            x = CDbl(points(i))
            sumSq = sumSq + (x - mean) ^ 2
        End If
    Next i

    ' Calcular desviación estándar
    CalculateStdDevOverall = Sqr(sumSq / (n - 1))
End Function

' Función para prueba de Shapiro-Wilk (placeholder - implementar o usar librería externa)
Private Function ShapiroWilkTest(points As Collection) As Boolean
    ' Para auditorías, implementar algoritmo completo o referenciar una fuente válida.
    ' Por ahora, retorna True asumiendo normalidad.
    ShapiroWilkTest = True
End Function

' Función para crear hoja de resultados
Private Function CreateResultsSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(sheetName).Delete
    Application.DisplayAlerts = True
    Set CreateResultsSheet = ActiveWorkbook.Sheets.Add
    CreateResultsSheet.Name = sheetName
End Function

' Procedimiento para generar histograma (esquema)
Private Sub GenerateHistogram(ws As Worksheet, points As Collection, _
    LSE As Double, LIE As Double, Target As Double)
    ' Nota: Esto requiere referenciar "Microsoft Excel Object Library" para gráficos.
    ' Implementar usando ChartObjects.Add y configurando series.
    ' Incluir líneas de LSE/LIE y Target en el gráfico.
End Sub

Function EstimateStdDevWithin(dataRange As Range) As Double
    Dim fila As Range
    Dim s As Double, sumaCuadrados As Double
    Dim k As Long
    Dim valores As Variant
    Dim i As Long, media As Double

    sumaCuadrados = 0
    k = dataRange.Rows.count

    For Each fila In dataRange.Rows
        valores = fila.Value
        media = 0
        For i = 1 To fila.Columns.count
            media = media + CDbl(valores(1, i))
        Next i
        media = media / fila.Columns.count

        s = 0
        For i = 1 To fila.Columns.count
            s = s + (CDbl(valores(1, i)) - media) ^ 2
        Next i
        s = s / (fila.Columns.count - 1)
        sumaCuadrados = sumaCuadrados + s
    Next fila

    If k > 0 Then
        EstimateStdDevWithin = Sqr(sumaCuadrados / k)
    Else
        EstimateStdDevWithin = 0
    End If
End Function

Public Function CreateUniqueSheet(baseName As String) As Worksheet
    Dim i As Integer
    Dim sheetName As String
    Dim wsCheck As Worksheet
    
    i = 1
    Do
        sheetName = baseName & "_" & i
        Set wsCheck = Nothing
        On Error Resume Next
        Set wsCheck = ActiveWorkbook.Sheets(sheetName)
        On Error GoTo 0
        i = i + 1
    Loop While Not wsCheck Is Nothing
    
    ' Crear la hoja con el nombre único
    Set CreateUniqueSheet = ActiveWorkbook.Sheets.Add
    CreateUniqueSheet.Name = sheetName
End Function



