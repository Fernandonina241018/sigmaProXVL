Option Explicit

'=====================
'  A) Lógica central
'=====================

'-------------------------------------------------------------------------------
' Función: CalcularEstadisticosPorColumna
' Descripción:
'   Calcula una serie de estadísticos básicos para cada columna de un rango
'   especificado. Retorna una matriz con encabezado, conteo, mediana, media,
'   varianza, desviación estándar, mínimo y máximo, por columna.
'
' Parámetros:
'   direccion          - Cadena que especifica el rango (por ejemplo, "[Libro1]Hoja1!$C$1:$F$5").
'   usarVarMuestral    - Booleano opcional, True para usar varianza muestral (VAR.S),
'                        False para varianza poblacional (VAR.P). Por defecto es True.
'
' Retorno:
'   Una matriz Variant (nCols x nStats) con los resultados calculados por columna.
'
' Errores:
'   Lanza error si el rango no puede resolverse o si no tiene al menos una fila de datos.
'-------------------------------------------------------------------------------
Public Function CalcularEstadisticosPorColumna( _
    ByVal direccion As String, _
    Optional ByVal usarVarMuestral As Boolean = True) As Variant

    Dim rng As Range, datosCol As Range
    Dim nCols As Long, nRows As Long
    Dim i As Long
    Dim cntNums As Long
    Dim header As Variant
    
    ' --- Estadísticos a calcular ---
    ' Orden de columnas de salida:
    ' [Encabezado, Conteo, Mediana, Media, Varianza, Desv.Est., Mínimo, Máximo]
    Const OUT_COLS As Long = 8
    Dim med As Variant, varx As Variant, avg As Variant, stdev As Variant
    Dim minv As Variant, maxv As Variant, cnt As Variant
    
    ' Obtener el rango desde el texto proporcionado
    Set rng = GetRangeFromAddress(direccion)
    If rng Is Nothing Then
        Err.Raise vbObjectError + 513, , "No se pudo resolver el rango desde la dirección proporcionada."
    End If
    
    If rng.Rows.Count < 2 Then
        Err.Raise vbObjectError + 514, , "El rango debe tener al menos 2 filas (encabezado + datos)."
    End If
    
    nCols = rng.Columns.Count
    nRows = rng.Rows.Count
    
    Dim res() As Variant
    ReDim res(1 To nCols, 1 To OUT_COLS)
    
    ' Recorrer columnas y calcular estadísticos
    For i = 1 To nCols
        ' Datos de la columna i (excluye el encabezado)
        Set datosCol = rng.Columns(i).Resize(nRows - 1, 1).Offset(1, 0)
        header = rng.Cells(1, i).Value
        
        ' Conteo de valores numéricos
        cntNums = Application.WorksheetFunction.Count(datosCol)
        cnt = cntNums
        
        ' Mediana
        If cntNums >= 1 Then
            med = Application.WorksheetFunction.Median(datosCol)
        Else
            med = CVErr(xlErrNA)
        End If
        
        ' Media (promedio)
        If cntNums >= 1 Then
            avg = Application.WorksheetFunction.Average(datosCol)
        Else
            avg = CVErr(xlErrNA)
        End If
        
        ' Varianza
        If usarVarMuestral Then
            If cntNums >= 2 Then
                On Error Resume Next
                varx = Application.WorksheetFunction.Var_S(datosCol)
                If Err.Number <> 0 Then
                    Err.Clear
                    varx = Application.WorksheetFunction.Var(datosCol) ' Compatibilidad versiones antiguas
                End If
                On Error GoTo 0
            Else
                varx = CVErr(xlErrDiv0)
            End If
        Else
            On Error Resume Next
            varx = Application.WorksheetFunction.Var_P(datosCol)
            If Err.Number <> 0 Then
                Err.Clear
                varx = Application.WorksheetFunction.VarP(datosCol) ' Compatibilidad versiones antiguas
            End If
            On Error GoTo 0
        End If
        
        ' Desviación estándar
        If usarVarMuestral Then
            If cntNums >= 2 Then
                On Error Resume Next
                stdev = Application.WorksheetFunction.StDev_S(datosCol)
                If Err.Number <> 0 Then
                    Err.Clear
                    stdev = Application.WorksheetFunction.StDev(datosCol) ' Compatibilidad versiones antiguas
                End If
                On Error GoTo 0
            Else
                stdev = CVErr(xlErrDiv0)
            End If
        Else
            On Error Resume Next
            stdev = Application.WorksheetFunction.StDev_P(datosCol)
            If Err.Number <> 0 Then
                Err.Clear
                stdev = Application.WorksheetFunction.StDevP(datosCol) ' Compatibilidad versiones antiguas
            End If
            On Error GoTo 0
        End If
        
        ' Mínimo y máximo
        If cntNums >= 1 Then
            minv = Application.WorksheetFunction.Min(datosCol)
            maxv = Application.WorksheetFunction.Max(datosCol)
        Else
            minv = CVErr(xlErrNA)
            maxv = CVErr(xlErrNA)
        End If
        
        ' Almacenar resultados en la matriz
        res(i, 1) = CStr(header)
        res(i, 2) = cnt
        res(i, 3) = med
        res(i, 4) = avg
        res(i, 5) = varx
        res(i, 6) = stdev
        res(i, 7) = minv
        res(i, 8) = maxv
    Next i
    
    CalcularEstadisticosPorColumna = res
End Function

'==========================
'  B) Resolver dirección
'==========================

'-------------------------------------------------------------------------------
' Función: GetRangeFromAddress
' Descripción:
'   Convierte una dirección de rango en formato texto (puede incluir libro y hoja)
'   en un objeto Range de Excel.
'
' Parámetros:
'   direccion - Texto con formato tipo "[Libro1]Hoja1!$C$1:$F$5" o "Hoja1!A1:B10"
'
' Retorno:
'   Range válido si la dirección puede ser procesada, Nothing en caso contrario.
'-------------------------------------------------------------------------------
Public Function GetRangeFromAddress(ByVal direccion As String) As Range
    Dim exclPos As Long
    Dim leftPart As String, addrPart As String
    Dim wbName As String, wsName As String
    Dim wb As Workbook, ws As Worksheet

    direccion = Trim(direccion)
    If Len(direccion) = 0 Then Exit Function

    exclPos = InStr(1, direccion, "!")
    If exclPos = 0 Then Exit Function

    leftPart = Left(direccion, exclPos - 1)   ' [Libro1]Hoja1  o  Hoja1
    addrPart = Mid(direccion, exclPos + 1)    ' $C$1:$F$5

    ' ¿Incluye nombre de libro entre corchetes?
    If InStr(1, leftPart, "[") > 0 And InStr(1, leftPart, "]") > 0 Then
        wbName = Mid(leftPart, InStr(1, leftPart, "[") + 1, _
                     InStr(1, leftPart, "]") - InStr(1, leftPart, "[") - 1)
        wsName = Mid(leftPart, InStr(1, leftPart, "]") + 1)
    Else
        wbName = ""          ' Asume libro activo
        wsName = leftPart
    End If
    On Error GoTo salir

    If wbName <> "" Then
        Set wb = Application.Workbooks(wbName)
        If wb Is Nothing Then GoTo salir
    Else
        Set wb = Application.ActiveWorkbook
    End If

    Set ws = wb.Worksheets(wsName)
    Set GetRangeFromAddress = ws.Range(addrPart)

    Exit Function
salir:
    Set GetRangeFromAddress = Nothing
End Function

'=========================================
'  C) Crear NUEVA hoja y escribir resultados
'=========================================

'-------------------------------------------------------------------------------
' Subrutina: EscribirResultadosEnNuevaHoja
' Descripción:
'   Genera una nueva hoja en el libro de Excel con los resultados estadísticos
'   calculados por columna. Si existe una hoja con el nombre base, se le agrega un sufijo.
'
' Parámetros:
'   direccion         - Dirección del rango de datos.
'   usarVarMuestral   - Booleano opcional para seleccionar tipo de varianza.
'   nombreBaseHoja    - Nombre base para la hoja a crear. Por defecto "Estadisticas".
'
' Efectos:
'   Crea hoja y vuelca los resultados y encabezados a partir de la celda A1.
'-------------------------------------------------------------------------------
Public Sub EscribirResultadosEnNuevaHoja( _
    ByVal direccion As String, _
    Optional ByVal usarVarMuestral As Boolean = True, _
    Optional ByVal nombreBaseHoja As String = "Estadisticas")

    Dim res As Variant
    Dim wsOut As Worksheet
    Dim headers As Variant
    Dim nFilas As Long, nCols As Long

    ' Calcular estadísticos
    res = CalcularEstadisticosPorColumna(direccion, usarVarMuestral)
    nFilas = UBound(res, 1)
    nCols = UBound(res, 2)

    ' Encabezados de columna
    headers = Array("Columna", "Conteo", "Mediana", "Media", _
                    IIf(usarVarMuestral, "Varianza (VAR.S)", "Varianza (VAR.P)"), _
                    IIf(usarVarMuestral, "Desv.Est. (STDEV.S)", "Desv.Est. (STDEV.P)"), _
                    "Mínimo", "Máximo")

    ' Crear hoja nueva con nombre disponible
    Set wsOut = CrearHojaConNombreDisponible(ThisWorkbook, nombreBaseHoja)

    ' Volcar encabezados y resultados
    wsOut.Range("A1").Resize(1, nCols).Value = headers
    wsOut.Range("A2").Resize(nFilas, nCols).Value = res
    ' Formato rápido: autofit y resaltado en primera fila
    With wsOut
        .Columns("A:" & Chr$(64 + nCols)).AutoFit
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(242, 242, 242)
    End With

    wsOut.Activate
End Sub

'=========================================
'  D) Utilidad: crear hoja con nombre disponible
'=========================================

'-------------------------------------------------------------------------------
' Función: CrearHojaConNombreDisponible
' Descripción:
'   Añade una hoja nueva con un nombre base, asegurando que no se repita en el libro.
'   Si el nombre existe, agrega un sufijo entre paréntesis incrementando.
'
' Parámetros:
'   wb          - Workbook donde se creará la hoja.
'   nombreBase  - Cadena base para el nombre de la hoja.
'
' Retorno:
'   Worksheet - Referencia a la hoja creada.
'-------------------------------------------------------------------------------
Public Function CrearHojaConNombreDisponible( _
    ByVal wb As Workbook, _
    ByVal nombreBase As String) As Worksheet

    Dim nombrePropuesto As String
    Dim idx As Long
    Dim existe As Boolean

    nombrePropuesto = nombreBase
    idx = 1

    ' Buscar un nombre disponible (agrega sufijos si necesario)
    Do
        existe = False
        Dim ws As Worksheet
        For Each ws In wb.Worksheets
            If StrComp(ws.Name, nombrePropuesto, vbTextCompare) = 0 Then
                existe = True
                Exit For
            End If
        Next ws

        If existe Then
            idx = idx + 1
            nombrePropuesto = nombreBase & " (" & idx & ")"
        Else
            Exit Do
        End If
    Loop

    Set CrearHojaConNombreDisponible = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    CrearHojaConNombreDisponible.Name = nombrePropuesto
End Function
