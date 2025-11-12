Sub IniciarMapaCalor()
    ' MACRO PRINCIPAL - Ejecutar esta para iniciar
    sigmaproxvl.Show
End Sub

Public Sub GenerarMapaCalorDesdeRangos(RangoDatos As Range, rangoDestino As Range, _
                                 Optional rangoFilas As Range, Optional rangoColumnas As Range, _
                                 Optional generarAleatorio As Boolean = False)

    Dim ws As Worksheet
    Dim filas As Long, columnas As Long
    Dim i As Long, j As Long
    Dim celdaActual As Range
    Dim offsetFila As Long, offsetCol As Long
    Dim objetivo As Double
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
    
    ' 1. Obtener el valor objetivo (Asumimos que el ComboBox está en un UserForm)
    ' Asegúrate de manejar el error 'Type Mismatch' si el ComboBox está vacío!
    If IsNumeric(sigmaproxvl.cboExpectativa.Value) Then
        objetivo = CDbl(sigmaproxvl.cboExpectativa.Value)
    Else
        MsgBox "El valor objetivo en el ComboBox no es numérico.", vbCritical
        Exit Sub
    End If

    Set ws = rangoDestino.Worksheet
    filas = RangoDatos.Rows.count
    columnas = RangoDatos.Columns.count

    Application.ScreenUpdating = False

    ' Limpiar área de destino
    ws.Range(rangoDestino, rangoDestino.Offset(filas + 10, columnas + 3)).Clear

    ' TÍTULO
    With rangoDestino
        .Value = "MAPA DE CALOR - ALMACÉN"
        .Font.Size = 14
        .Font.Bold = True
        .Font.color = RGB(31, 78, 121)
    End With

    ' Instrucciones
    With rangoDestino.Offset(1, 0)
        .Value = "Valores: 0-100 | Colores: Azul(bajo) ? Verde ? Amarillo ? Naranja ? Rojo(alto)"
        .Font.Size = 9
        .Font.Italic = True
    End With

    offsetFila = 3
    offsetCol = 1

    ' ENCABEZADOS DE COLUMNAS
    For j = 1 To columnas
        Set celdaActual = rangoDestino.Offset(offsetFila, offsetCol + j - 1)

        If Not rangoColumnas Is Nothing And j <= rangoColumnas.Cells.count Then
            celdaActual.Value = rangoColumnas.Cells(j).Value
        Else
            celdaActual.Value = "Z" & j
        End If

        With celdaActual
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.color = RGB(220, 230, 241)
            .Font.Size = 10
        End With
    Next j

    ' ENCABEZADOS DE FILAS
    For i = 1 To filas
        Set celdaActual = rangoDestino.Offset(offsetFila + i, 0)

        If Not rangoFilas Is Nothing And i <= rangoFilas.Cells.count Then
            celdaActual.Value = rangoFilas.Cells(i).Value
        Else
            celdaActual.Value = "P" & i
        End If

        With celdaActual
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.color = RGB(220, 230, 241)
            .Font.Size = 10
        End With
    Next i

    ' ÁREA DE DATOS DEL MAPA
    Dim rangoMapa As Range
    Set rangoMapa = ws.Range(rangoDestino.Offset(offsetFila + 1, offsetCol), _
                             rangoDestino.Offset(offsetFila + filas, offsetCol + columnas - 1))

    ' Copiar valores o generar aleatorios
    If generarAleatorio Then
        Dim c As Range
        For Each c In rangoMapa
            c.Value = Int(Rnd() * 101)
        Next c
    Else
        rangoMapa.Value = RangoDatos.Value
    End If

    ' FORMATEAR CELDAS DEL MAPA
    With rangoMapa
        .ColumnWidth = 10
        .RowHeight = 35
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 12
        .Borders.Weight = xlThin
        .Borders.color = RGB(200, 200, 200)
    End With

    ' APLICAR COLORES
    For Each celdaActual In rangoMapa
        Call AplicarColorDesviacion(celdaActual, objetivo)
    Next celdaActual

    ' CREAR LEYENDA
    Dim filaLeyenda As Long
    filaLeyenda = offsetFila + filas + 3

    With rangoDestino.Offset(filaLeyenda, 0)
        .Value = "LEYENDA:"
        .Font.Bold = True
        .Font.Size = 11
    End With

    Dim leyendas() As Variant
    Dim colores() As Variant

    leyendas = Array("0-19: Muy Bajo", "20-39: Bajo", "40-59: Medio", _
                     "60-79: Alto", "80-89: Muy Alto", "90-100: Crítico")
    colores = Array(RGB(30, 58, 138), RGB(59, 130, 246), RGB(34, 197, 94), _
                    RGB(251, 191, 36), RGB(249, 115, 22), RGB(220, 38, 38))

    For i = 0 To 5
        With rangoDestino.Offset(filaLeyenda + i, 1)
            .Interior.color = colores(i)
            .ColumnWidth = 3
            .Borders.Weight = xlThin
        End With

        With rangoDestino.Offset(filaLeyenda + i, 2)
            .Value = leyendas(i)
            .Font.Size = 9
        End With
    Next i

    ' ESTADÍSTICAS
    Dim estadFila As Long
    estadFila = filaLeyenda

    With rangoDestino.Offset(estadFila, 5)
        .Value = "ESTADÍSTICAS:"
        .Font.Bold = True
        .Font.Size = 11
    End With

    With rangoDestino.Offset(estadFila + 1, 5)
        .Value = "Promedio:"
        .Offset(0, 1).Value = Application.WorksheetFunction.Average(rangoMapa)
        .Offset(0, 1).NumberFormat = "0.00"
    End With

    With rangoDestino.Offset(estadFila + 2, 5)
        .Value = "Máximo:"
        .Offset(0, 1).Value = Application.WorksheetFunction.Max(rangoMapa)
    End With

    With rangoDestino.Offset(estadFila + 3, 5)
        .Value = "Mínimo:"
        .Offset(0, 1).Value = Application.WorksheetFunction.Min(rangoMapa)
    End With

    ' GUARDAR REFERENCIA DEL RANGO
    On Error Resume Next
    ws.Names.Add Name:="RangoMapaCalor", RefersTo:=rangoMapa
    On Error GoTo 0

    Application.ScreenUpdating = True

    Debug.Print "¡Mapa de calor generado exitosamente!" & vbCrLf & vbCrLf & _
           "Dimensiones: " & filas & " filas x " & columnas & " columnas" & vbCrLf & _
           "Total de celdas: " & filas * columnas & vbCrLf & vbCrLf & _
           "Use 'ActualizarMapaCalor' para refrescar los colores.", _
           vbInformation, "Completado"
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
End Sub

Public Sub ActualizarMapaCalor()
    ' Actualiza los colores del mapa existente después de cambiar valores
    Dim rangoMapa As Range
    Dim celda As Range
    
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

    On Error Resume Next
    Set rangoMapa = ActiveSheet.Range("RangoMapaCalor")
    On Error GoTo 0

    If rangoMapa Is Nothing Then
        MsgBox "No se encontró un mapa de calor en esta hoja." & vbCrLf & vbCrLf & _
               "Primero debe generar un mapa usando la macro 'IniciarMapaCalor'.", _
               vbExclamation, "Error"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    For Each celda In rangoMapa
        Call AplicarColorDesviacion(celdaActual, objetivo)
    Next celda

    Application.ScreenUpdating = True

    MsgBox "Mapa de calor actualizado correctamente.", vbInformation, "Actualizado"
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
End Sub

Sub LimpiarMapaCalor()
    ' Limpia todos los valores del mapa (los pone en 0)
    Dim rangoMapa As Range
    Dim respuesta As VbMsgBoxResult
    
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

    On Error Resume Next
    Set rangoMapa = ActiveSheet.Range("RangoMapaCalor")
    On Error GoTo 0

    If rangoMapa Is Nothing Then
        MsgBox "No se encontró un mapa de calor en esta hoja.", vbExclamation, "Error"
        Exit Sub
    End If

    respuesta = MsgBox("¿Está seguro que desea limpiar todos los valores del mapa?" & vbCrLf & _
                       "Todos los datos se establecerán en 0.", _
                       vbQuestion + vbYesNo, "Confirmar")

    If respuesta = vbYes Then
        Application.ScreenUpdating = False
        rangoMapa.Value = 0
        Call ActualizarMapaCalor
        Application.ScreenUpdating = True
        MsgBox "Mapa limpiado exitosamente.", vbInformation
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

Sub ExportarMapaCSV()
    ' Exporta el mapa de calor a un archivo CSV
    Dim rangoMapa As Range
    Dim filePath As String
    Dim fileNum As Integer
    Dim i As Long, j As Long
    Dim linea As String
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

    On Error Resume Next
    Set rangoMapa = ActiveSheet.Range("RangoMapaCalor")
    On Error GoTo 0

    If rangoMapa Is Nothing Then
        MsgBox "No hay mapa de calor para exportar en esta hoja.", vbExclamation, "Error"
        Exit Sub
    End If

    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="MapaCalor_" & Format(Now, "yyyymmdd_hhmmss") & ".csv", _
        FileFilter:="Archivos CSV (*.csv), *.csv", _
        Title:="Guardar Mapa de Calor")

    If filePath = "False" Then Exit Sub

    fileNum = FreeFile
    Open filePath For Output As #fileNum

    ' Escribir datos
    For i = 1 To rangoMapa.Rows.count
        linea = ""
        For j = 1 To rangoMapa.Columns.count
            linea = linea & rangoMapa.Cells(i, j).Value
            If j < rangoMapa.Columns.count Then linea = linea & ","
        Next j
        Print #fileNum, linea
    Next i

    Close #fileNum

    MsgBox "Datos exportados exitosamente a:" & vbCrLf & vbCrLf & filePath, _
           vbInformation, "Exportación Completa"
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
    
End Sub

Sub ImportarDatosCSV()
    ' Importa datos desde un archivo CSV al mapa existente
    Dim rangoMapa As Range
    Dim filePath As Variant
    Dim fileNum As Integer
    Dim linea As String
    Dim datos() As String
    Dim i As Long, j As Long
    Dim fila As Long
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

    On Error Resume Next
    Set rangoMapa = ActiveSheet.Range("RangoMapaCalor")
    On Error GoTo 0

    If rangoMapa Is Nothing Then
        MsgBox "Primero debe crear un mapa de calor.", vbExclamation, "Error"
        Exit Sub
    End If

    filePath = Application.GetOpenFilename( _
        FileFilter:="Archivos CSV (*.csv), *.csv", _
        Title:="Seleccionar archivo CSV")

    If filePath = False Then Exit Sub

    fileNum = FreeFile
    Open filePath For Input As #fileNum

    fila = 1
    Do While Not EOF(fileNum) And fila <= rangoMapa.Rows.count
        Line Input #fileNum, linea
        datos = Split(linea, ",")

        For j = 0 To UBound(datos)
            If j < rangoMapa.Columns.count Then
                rangoMapa.Cells(fila, j + 1).Value = Val(datos(j))
            End If
        Next j

        fila = fila + 1
    Loop

    Close #fileNum

    Call ActualizarMapaCalor

    MsgBox "Datos importados y mapa actualizado exitosamente.", vbInformation
    
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
    
End Sub

