Public Sub CargarRangoEnListBox(lst As MSForms.ListBox, rangoTexto As String)
    Dim rango As Range
    Dim datos As Variant
    Dim i As Long, j As Long
    Dim fila As Long
    Dim numColumnas As Long
    Dim numFilas As Long
    
    On Error GoTo ErrorHandler
    
    ' ===== PASO 1: OBTENER EL RANGO =====
    Set rango = ObtenerRangoDesdeTexto(rangoTexto)
    
    If rango Is Nothing Then
        MsgBox "No se pudo acceder al rango especificado:" & vbCrLf & vbCrLf & _
               rangoTexto, vbCritical, "Error de Rango"
        Exit Sub
    End If
    
    ' ===== PASO 2: LEER DATOS DEL RANGO =====
    datos = rango.Value
    
    ' Verificar si es un array bidimensional
    If Not IsArray(datos) Then
        ' Si es una sola celda, convertir a array
        ReDim datos(1 To 1, 1 To 1)
        datos(1, 1) = rango.Value
    End If
    
    ' ===== PASO 3: DETERMINAR DIMENSIONES =====
    If IsArray(datos) Then
        On Error Resume Next
        numFilas = UBound(datos, 1) - LBound(datos, 1) + 1
        numColumnas = UBound(datos, 2) - LBound(datos, 2) + 1
        On Error GoTo ErrorHandler
    Else
        numFilas = 1
        numColumnas = 1
    End If
    
    ' ===== PASO 4: CONFIGURAR LISTBOX =====
    With lst
        .Clear
        .ColumnCount = numColumnas
        
        ' Ajustar ancho de columnas automáticamente
        .ColumnWidths = GenerarAnchoColumnas(numColumnas, lst.Width)
    End With
    
    ' ===== PASO 5: CARGAR DATOS EN LISTBOX =====
    For i = LBound(datos, 1) To UBound(datos, 1)
        ' Agregar fila
        lst.AddItem datos(i, LBound(datos, 2))
    Next i
    
    ' Mensaje de confirmación
    Debug.Print "Datos cargados exitosamente:" & vbCrLf & vbCrLf & _
           "Filas: " & numFilas & vbCrLf & _
           "Columnas: " & numColumnas, _
           vbInformation, "Carga Completa"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al cargar datos en el ListBox:" & vbCrLf & vbCrLf & _
           "Número: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, _
           vbCritical, "Error"
End Sub

' =============================================================================
' FUNCIÓN AUXILIAR: Obtener Rango desde Texto
' =============================================================================
Private Function ObtenerRangoDesdeTexto(rangoTexto As String) As Range
    ' Maneja referencias tanto locales como externas
    ' Ejemplos:
    ' - "A1:B10"
    ' - "Hoja1!A1:B10"
    ' - "'Hoja con espacios'!A1:B10"
    ' - "[Archivo.xlsx]Hoja1!A1:B10"
    ' - "'[4A. DATOS.XLSX]PRINCIPIO ACTIVO'!$D$4:$E$10"
    
    Dim rng As Range
    
    On Error Resume Next
    
    ' Intentar obtener el rango
    Set rng = Application.Range(rangoTexto)
    
    ' Si falla, intentar con Evaluate
    If rng Is Nothing Then
        Set rng = Application.Evaluate(rangoTexto)
    End If
    
    On Error GoTo 0
    
    Set ObtenerRangoDesdeTexto = rng
    
End Function

' =============================================================================
' FUNCIÓN AUXILIAR: Generar Ancho de Columnas
' =============================================================================
Private Function GenerarAnchoColumnas(numColumnas As Long, anchoTotal As Single) As String
    ' Distribuye el ancho equitativamente entre columnas
    
    Dim anchoColumna As Single
    Dim i As Long
    Dim resultado As String
    
    ' Calcular ancho por columna (restando un margen)
    anchoColumna = (anchoTotal - 20) / numColumnas
    
    ' Generar string de anchos
    For i = 1 To numColumnas
        If i = 1 Then
            resultado = CStr(anchoColumna)
        Else
            resultado = resultado & ";" & CStr(anchoColumna)
        End If
    Next i
    
    GenerarAnchoColumnas = resultado
    
End Function


