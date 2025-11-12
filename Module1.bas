Option Explicit

'=====================
'  A) Lógica central
'=====================

' Devuelve una matriz (nCols x nStats) con los estadísticos por columna.
' dirección: por ejemplo "[Book1]Sheet1!$C$1:$F$5" (leída desde TextBox).
' usarVarMuestral: True = VAR.S (muestral), False = VAR.P (poblacional).
' Inicia análisis después del encabezado (fila + 1).
Public Function CalcularEstadisticosPorColumna( _
    ByVal direccion As String, _
    Optional ByVal usarVarMuestral As Boolean = True) As Variant

    Dim rng As Range, datosCol As Range
    Dim nCols As Long, nRows As Long
    Dim i As Long
    Dim cntNums As Long
    Dim header As Variant
    
    ' --- Estadísticos a calcular (ajusta aquí si quieres más/menos) ---
    ' Orden de columnas de salida:
    ' [Encabezado, Conteo, Mediana, Media, Varianza, Desv.Est., Mínimo, Máximo]
    Const OUT_COLS As Long = 8
    Dim med As Variant, varx As Variant, avg As Variant, stdev As Variant
    Dim minv As Variant, maxv As Variant, cnt As Variant
    
    ' Resolver el rango desde el texto
    Set rng = GetRangeFromAddress(direccion)
    If rng Is Nothing Then
        Err.Raise vbObjectError + 513, , "No se pudo resolver el rango desde la dirección proporcionada."
    End If
    
    If rng.Rows.count < 2 Then
        Err.Raise vbObjectError + 514, , "El rango debe tener al menos 2 filas (encabezado + datos)."
    End If
    
    nCols = rng.Columns.count
    nRows = rng.Rows.count
    
    Dim res() As Variant
    ReDim res(1 To nCols, 1 To OUT_COLS)
    
    For i = 1 To nCols
        ' Datos de la columna i (excluye encabezado)
        Set datosCol = rng.Columns(i).Resize(nRows - 1, 1).Offset(1, 0)
        
        header = rng.Cells(1, i).Value
        
        ' Conteo de valores numéricos
        cntNums = Application.WorksheetFunction.count(datosCol)
        cnt = cntNums
        
        ' Mediana
        If cntNums >= 1 Then
            med = Application.WorksheetFunction.Median(datosCol)
        Else
            med = CVErr(xlErrNA)
        End If
        
        ' Media
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
                    varx = Application.WorksheetFunction.Var(datosCol) ' compatibilidad
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
                varx = Application.WorksheetFunction.VarP(datosCol) ' compatibilidad
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
                    stdev = Application.WorksheetFunction.stdev(datosCol) ' compatibilidad
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
                stdev = Application.WorksheetFunction.StDevP(datosCol) ' compatibilidad
            End If
            On Error GoTo 0
        End If
        
        ' Mínimo y Máximo
        If cntNums >= 1 Then
            minv = Application.WorksheetFunction.Min(datosCol)
            maxv = Application.WorksheetFunction.Max(datosCol)
        Else
            minv = CVErr(xlErrNA)
            maxv = CVErr(xlErrNA)
        End If
        
        ' Cargar resultados en matriz
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

' Convierte "[Book1]Sheet1!$C$1:$F$5" o "Sheet1!A1:B10" a Range
Public Function GetRangeFromAddress(ByVal direccion As String) As Range
    Dim exclPos As Long
    Dim leftPart As String, addrPart As String
    Dim wbName As String, wsName As String
    Dim wb As Workbook, ws As Worksheet
    direccion = Trim(direccion)
    If Len(direccion) = 0 Then Exit Function

    exclPos = InStr(1, direccion, "!")
    If exclPos = 0 Then Exit Function

    leftPart = Left(direccion, exclPos - 1)   ' [Book1]Sheet1  o  Sheet1
    addrPart = Mid(direccion, exclPos + 1)    ' $C$1:$F$5

    ' ¿Incluye nombre de libro entre corchetes?
    If InStr(1, leftPart, "[") > 0 And InStr(1, leftPart, "]") > 0 Then
        wbName = Mid(leftPart, InStr(1, leftPart, "[") + 1, _
                     InStr(1, leftPart, "]") - InStr(1, leftPart, "[") - 1)
        wsName = Mid(leftPart, InStr(1, leftPart, "]") + 1)
    Else
        wbName = ""          ' asume libro activo
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

' Crea una nueva hoja con nombre base "Estadisticas" (si existe, agrega sufijo).
' Escribe los resultados comenzando en A1 de esa hoja.
Public Sub EscribirResultadosEnNuevaHoja( _
    ByVal direccion As String, _
    Optional ByVal usarVarMuestral As Boolean = True, _
    Optional ByVal nombreBaseHoja As String = "Estadisticas")

    Dim res As Variant
    Dim wsOut As Worksheet
    Dim headers As Variant
    Dim nFilas As Long, nCols As Long

    ' Calcular
    res = CalcularEstadisticosPorColumna(direccion, usarVarMuestral)
    nFilas = UBound(res, 1)
    nCols = UBound(res, 2)

    ' Encabezados (ajustar si agregas o quitas estadísticas)
    headers = Array("Columna", "Conteo", "Mediana", "Media", _
                    IIf(usarVarMuestral, "Varianza (VAR.S)", "Varianza (VAR.P)"), _
                    IIf(usarVarMuestral, "Desv.Est. (STDEV.S)", "Desv.Est. (STDEV.P)"), _
                    "Mínimo", "Máximo")

    ' Crear hoja nueva con nombre disponible
    Set wsOut = CrearHojaConNombreDisponible(ThisWorkbook, nombreBaseHoja)

    ' Volcar encabezados y datos
    wsOut.Range("A1").Resize(1, nCols).Value = headers
    wsOut.Range("A2").Resize(nFilas, nCols).Value = res
    ' Formato rápido
    With wsOut
        .Columns("A:" & Chr$(64 + nCols)).AutoFit
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.color = RGB(242, 242, 242)
    End With

    wsOut.Activate
End Sub


'=========================================
'  D) Utilidad: crear hoja con nombre disponible
'=========================================
Public Function CrearHojaConNombreDisponible( _
    ByVal wb As Workbook, _
    ByVal nombreBase As String) As Worksheet

    Dim nombrePropuesto As String
    Dim idx As Long
    Dim existe As Boolean

    nombrePropuesto = nombreBase
    idx = 1

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

    Set CrearHojaConNombreDisponible = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    CrearHojaConNombreDisponible.Name = nombrePropuesto
End Function

