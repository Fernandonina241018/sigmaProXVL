Option Explicit

' =====================================================
' FUNCIONES DE FORMATO
' =====================================================

Public Sub FormatearEncabezadoPrincipal(rng As Range)
    With rng
        .Font.Bold = True
        .Font.Size = 14
        .Font.color = RGB(0, 51, 102)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.color = RGB(220, 230, 241)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Merge
    End With
End Sub

Public Sub FormatearEncabezadoSeccion(rng As Range)
    With rng
        .Font.Bold = True
        .Font.Size = 12
        .Font.color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Interior.color = RGB(0, 51, 102)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Merge
    End With
End Sub

Sub AjustarAlturaPorTextoCombinado(celda As Range)
    Dim texto As String
    Dim largo As Long
    Dim lineasEstimadas As Long
    Dim alturaBase As Double

    texto = celda.Value
    largo = Len(texto)
    alturaBase = 15 ' Altura mínima para una línea

    ' Estimar número de líneas (cada 40 caracteres ˜ 1 línea)
    lineasEstimadas = WorksheetFunction.RoundUp(largo / 40, 0)

    ' Ajustar altura proporcionalmente
    celda.Rows.RowHeight = alturaBase + (lineasEstimadas - 1) * 15
End Sub


Sub AjustarAlturaPorContenido(celda As Range)
    Dim texto As String
    Dim largo As Long
    Dim alturaBase As Double

    texto = celda.Value
    largo = Len(texto)
    alturaBase = 15 ' Altura mínima

    ' Estimación: cada 40 caracteres ˜ 1 línea adicional
    Dim lineasEstimadas As Long
    lineasEstimadas = WorksheetFunction.RoundUp(largo / 40, 0)

    ' Ajustar altura proporcionalmente
    celda.Rows.RowHeight = alturaBase + (lineasEstimadas - 1) * 15
End Sub

Sub FormatearDato(celda As Range)
    With celda
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = False
        .Font.color = RGB(0, 0, 0) ' Negro
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.color = RGB(255, 255, 255)
        
        ' Bordes
        Dim b As Variant
        For Each b In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
            .Borders(b).LineStyle = xlContinuous
            .Borders(b).Weight = xlThin
        Next b

        ' Formato según tipo
        If IsDate(.Value) Then
            .NumberFormat = "dd/mmm/yyyy hh:mm:ss AM/PM"
        ElseIf IsNumeric(.Value) Then
            .NumberFormat = "0.0000"
        Else
            .NumberFormat = "@"
        End If
    End With
End Sub


Sub FormatearEncabezado(rng As Range)
    With rng
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.color = RGB(0, 0, 0)
        .Interior.color = RGB(200, 200, 200)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
    End With
End Sub

Sub FormatearBloqueDatos(rng As Range)
    Dim celda As Range
    For Each celda In rng.Cells
        With celda
            .Font.Name = "Segoe UI"
            .Font.Size = 12
            .Font.Bold = True
            .Font.color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.color = RGB(0, 102, 204)

            ' Bordes continuos
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeBottom).Weight = xlThin
            .Borders(xlEdgeRight).Weight = xlThin

            ' Formato según tipo de dato
            If IsDate(.Value) Then
                .NumberFormat = "dd/mm/yyyy hh:mm:ss AM/PM"
            ElseIf IsNumeric(.Value) Then
                .NumberFormat = "#,##0.00"
            Else
                .NumberFormat = "@"
            End If
        End With
    Next celda
End Sub


Public Sub FormatearCeldaDatos(rng As Range)
    With rng
        ' Fuente profesional
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = False
        .Font.color = RGB(0, 0, 0) ' Negro

        ' Alineación
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        ' Fondo contrastante
        .Interior.color = RGB(255, 255, 255) ' Blanco
        
        ' Bordes continuos en todos los lados
        Dim b As Variant
        For Each b In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
            .Borders(b).LineStyle = xlContinuous
            .Borders(b).Weight = xlThin
        Next b

        ' Formato según tipo de dato
        If IsDate(.Value) Then
            .NumberFormat = "dd/mmm/yyyy hh:mm:ss AM/PM"
        ElseIf IsNumeric(.Value) Then
            .NumberFormat = "#,##0.0000"
        Else
            .NumberFormat = "@" ' texto
        End If


        ' Combinar si es un bloque
        If .Cells.count > 1 Then .Merge
    End With
End Sub

Public Sub FormatearEncabezadoTabla(rng As Range)
    With rng
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.color = RGB(200, 200, 200)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
End Sub

Public Function FormatoDeHoraUniversal(ByVal InputDateTime As Date) As String
' Propósito: Formatea un valor de fecha/hora VBA (Date) a una cadena de texto estándar.
' Argumento:
'   InputDateTime: El valor de fecha/hora que se desea formatear (usa ByVal para proteger el original).
' Devuelve: Una cadena de texto con el formato "DD/MMM/AAAA HH:MM:SS AM/PM".

    Const FORMATO_ESTANDAR As String = "dd/mmm/yyyy hh:mm:ss AM/PM"
    
    ' La función Format() convierte el valor Date en una cadena de texto (String)
    FormatoDeHoraUniversal = Format(InputDateTime, FORMATO_ESTANDAR)

End Function
