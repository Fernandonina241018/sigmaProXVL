' =============================================================================
' SELECTOR DE RANGO MEJORADO - SIMULA REFEDIT CON TEXTBOX + BOTÓN
' =============================================================================
' Autor: Optimizado
' Fecha: 2025
' Propósito: Permitir seleccionar rangos de Excel desde un UserForm usando TextBox
' =============================================================================

' =============================================================================
' VERSIÓN MEJORADA - CON VALIDACIONES Y MANEJO DE ERRORES
' =============================================================================
Sub ObtenerRango(ByVal TextBoxComoRefEdit As MSForms.TextBox)
    
    Dim RangoSeleccionado As Range
    Dim FormularioActivo As Object
    
    ' Guardar referencia al formulario activo
    Set FormularioActivo = sigmaproxvl ' Cambiar por el nombre de tu UserForm
    
    ' Ocultar temporalmente el formulario para mejor visibilidad
    FormularioActivo.Hide
    
    On Error Resume Next
    
    ' Solicitar selección de rango
    Set RangoSeleccionado = Application.InputBox( _
        Prompt:="Selecciona el rango de datos en la hoja de Excel:" & vbCrLf & vbCrLf & _
                "• Puedes seleccionar celdas individuales o rangos" & vbCrLf & _
                "• Usa Ctrl para selecciones múltiples" & vbCrLf & _
                "• Presiona ESC para cancelar", _
        Title:="Selector de Rango", _
        Default:=TextBoxComoRefEdit.Value, _
        Type:=8)
    
    On Error GoTo 0
    
    ' Mostrar nuevamente el formulario
    FormularioActivo.Show vbModeless
    
    ' Validar selección
    If Not RangoSeleccionado Is Nothing Then
        ' Actualizar TextBox con la dirección completa del rango
        TextBoxComoRefEdit.Value = RangoSeleccionado.Address(External:=True)
        TextBoxComoRefEdit.Tag = RangoSeleccionado.Address(False, False) ' Guardar dirección relativa
    Else
        ' Usuario canceló - mantener valor anterior o limpiar
         TextBoxComoRefEdit.Value = "" ' Descomenta para limpiar al cancelar
    End If
    
End Sub

' =============================================================================
' VERSIÓN ALTERNATIVA - CON OPCIONES AVANZADAS
' =============================================================================
Sub ObtenerRangoAvanzado(ByVal TextBoxComoRefEdit As MSForms.TextBox, _
                         Optional ByVal PermitirMultipleAreas As Boolean = False, _
                         Optional ByVal MostrarNombreHoja As Boolean = True)
    
    Dim RangoSeleccionado As Range
    Dim DireccionRango As String
    
    ' Ocultar formulario temporalmente
    Me.Hide
    
    On Error Resume Next
    Set RangoSeleccionado = Application.InputBox( _
        Prompt:="Selecciona el rango de datos:", _
        Title:="Selector de Rango", _
        Default:=IIf(TextBoxComoRefEdit.Value <> "", TextBoxComoRefEdit.Value, ActiveCell.Address), _
        Type:=8)
    On Error GoTo 0
    
    ' Restaurar formulario
    Me.Show vbModeless
    
    ' Procesar selección
    If Not RangoSeleccionado Is Nothing Then
        
        ' Validar si se permiten múltiples áreas
        If Not PermitirMultipleAreas And RangoSeleccionado.Areas.count > 1 Then
            MsgBox "Solo se permite seleccionar un rango continuo." & vbCrLf & _
                   "No uses Ctrl para selecciones múltiples.", _
                   vbExclamation, "Selección Inválida"
            Exit Sub
        End If
        
        ' Determinar formato de dirección
        If MostrarNombreHoja Then
            DireccionRango = RangoSeleccionado.Address(External:=True)
        Else
            DireccionRango = RangoSeleccionado.Address(False, False)
        End If
        
        ' Actualizar TextBox
        TextBoxComoRefEdit.Value = DireccionRango
        TextBoxComoRefEdit.Tag = RangoSeleccionado.Address ' Guardar referencia
        
    End If
    
End Sub

' =============================================================================
' FUNCIÓN AUXILIAR - VALIDAR RANGO DESDE TEXTBOX
' =============================================================================
Function ObtenerRangoDesdeTexto(ByVal TextBoxComoRefEdit As MSForms.TextBox) As Range
    
    Dim RangoResultado As Range
    
    On Error Resume Next
    
    ' Intentar obtener el rango desde el texto del TextBox
    If TextBoxComoRefEdit.Value <> "" Then
        Set RangoResultado = Range(TextBoxComoRefEdit.Value)
        
        ' Si falla, intentar con la dirección completa
        If RangoResultado Is Nothing Then
            Set RangoResultado = Application.Range(TextBoxComoRefEdit.Value)
        End If
    End If
    
    On Error GoTo 0
    
    ' Validar que se obtuvo un rango válido
    If RangoResultado Is Nothing Then
        MsgBox "El rango especificado no es válido:" & vbCrLf & vbCrLf & _
               TextBoxComoRefEdit.Value, vbExclamation, "Rango Inválido"
    End If
    
    Set ObtenerRangoDesdeTexto = RangoResultado
    
End Function

' =============================================================================
' FUNCIÓN AUXILIAR - VALIDAR SI UN RANGO ES VÁLIDO
' =============================================================================
Function ValidarRango(ByVal DireccionRango As String) As Boolean
    
    Dim RangoTemp As Range
    
    On Error Resume Next
    Set RangoTemp = Range(DireccionRango)
    On Error GoTo 0
    
    ValidarRango = Not RangoTemp Is Nothing
    
End Function

