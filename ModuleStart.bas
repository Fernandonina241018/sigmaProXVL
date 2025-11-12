Public Sub StartInteraccion()
    ' Mostrar formulario de forma MODAL (bloquea ejecución hasta cerrar)
    sigmaproxvl.Show
    
End Sub

Public Sub IniciarRegresionConRefEdit()
    Dim rngX As Range  ' Rango para la variable Independiente (X)
    Dim rngY As Range  ' Rango para la variable Dependiente (Y)
    
    On Error GoTo CancelHandler
    ' 1. Solicitar el Rango X (Selección Gráfica)
    'Set rngX = Application.Range(sigmaproxvl.RefEditVariableX.Text)
    
        ' 1. Seleccionar el Rango X
    Set rngX = Application.Range(sigmaproxvl.txtVariableX.Value)
        
    ' 2. Validación: Asegurar que la selección no sea bidimensional (X debe ser una columna)
    If rngX.Columns.count > 1 Then
        MsgBox "Error: Por favor, selecciona solo UNA columna para la variable X (independiente).", vbCritical
        Exit Sub
    End If
    
    ' 3. Derivar el Rango Y (Columna adyacente)
    Set rngY = rngX.Offset(0, 1)
    
    ' 4. Ejecutar la Regresión Lineal
    ' La llamada a la función ahora usa los rangos definidos dinámicamente.
    Call RegresionLinealSimpleMejorada(rngX, rngY)
    
    MsgBox "Análisis de regresión simple iniciado exitosamente con rangos seleccionados.", vbInformation
    
    Exit Sub
    
CancelHandler:
    If Err.Number = 424 Or Err.Number = 0 Then
        ' El usuario pulsó 'Cancelar'
    Else
        MsgBox "Se produjo un error: " & Err.Description, vbCritical
    End If
    
End Sub
