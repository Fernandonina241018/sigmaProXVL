Public Sub PopulateResultsSheet(ws As Worksheet, n As Long, mean As Double, _
    stdDevOverall As Double, stdDevWithin As Double, Cp As Double, Cpk As Double, _
    Pp As Double, Ppk As Double, Cpm As Double, IsNormal As Boolean, _
    LSE As Double, LIE As Double, Target As Double)
    
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
    
    With ws
        .Range("A1").Value = "Análisis de Capacidad del Proceso"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1:B1").Merge
        .Range("A1:B1").Interior.color = RGB(200, 200, 200) ' Fondo gris claro
        
        ' Sección 1: Estadísticas Básicas
        .Range("A3").Value = "Estadísticas Básicas"
        .Range("A3").Font.Bold = True
        .Range("A3").Font.Size = 12
        .Range("A3:B3").Merge
        .Range("A3:B3").Interior.color = RGB(200, 200, 200) ' Fondo gris claro
        
        .Range("A4").Value = "Tamaño de muestra (n):"
        .Range("A4").Font.Bold = True
        .Range("B4").Value = n
        .Range("B4").NumberFormat = "0"
        ' Aplica la alineación horizontal: Centro
        .Range("B4").HorizontalAlignment = xlCenter
        ' Aplica la alineación vertical (opcional, pero buena práctica)
        .Range("B4").VerticalAlignment = xlCenter
        
        .Range("A5").Value = "Media:"
        .Range("A5").Font.Bold = True
        .Range("B5").Value = mean
        .Range("B5").NumberFormat = "0.0000"
        ' Aplica la alineación horizontal: Centro
        .Range("B5").HorizontalAlignment = xlCenter
        ' Aplica la alineación vertical (opcional, pero buena práctica)
        .Range("B5").VerticalAlignment = xlCenter
        
        .Range("A6").Value = "Desv. Estándar (Overall):"
        .Range("A6").Font.Bold = True
        .Range("B6").Value = stdDevOverall
        .Range("B6").NumberFormat = "0.0000"
        ' Aplica la alineación horizontal: Centro
        .Range("B6").HorizontalAlignment = xlCenter
        ' Aplica la alineación vertical (opcional, pero buena práctica)
        .Range("B6").VerticalAlignment = xlCenter
        
        .Range("A7").Value = "Desv. Estándar (Within):"
        .Range("A7").Font.Bold = True
        .Range("B7").Value = stdDevWithin
        .Range("B7").NumberFormat = "0.0000"
        ' Aplica la alineación horizontal: Centro
        .Range("B7").HorizontalAlignment = xlCenter
        ' Aplica la alineación vertical (opcional, pero buena práctica)
        .Range("B7").VerticalAlignment = xlCenter
        
        ' Sección 2: Índices de Capacidad
        .Range("A9").Value = "Índices de Capacidad"
        .Range("A9").Font.Bold = True
        .Range("A9").Font.Size = 12
        .Range("A9:B9").Merge
        .Range("A9:B9").Interior.color = RGB(200, 200, 200)
        
        .Range("A10").Value = "Cp [ Capacidad potencial del proceso ]"
        .Range("A10").Font.Bold = True
        .Range("B10").Value = Cp
        .Range("B10").NumberFormat = "0.0000"
        ' Aplica la alineación horizontal: Centro
        .Range("B10").HorizontalAlignment = xlCenter
        ' Aplica la alineación vertical (opcional, pero buena práctica)
        .Range("B10").VerticalAlignment = xlCenter
        
        ' Formato condicional para Cp: rojo si < 1.33, amarillo si < 1.67, verde si >= 1.67
        If Cp < 1.33 Then
            .Range("B10").Interior.color = RGB(255, 0, 0) ' Rojo
        ElseIf Cp < 1.67 Then
            .Range("B10").Interior.color = RGB(255, 255, 0) ' Amarillo
        Else
            .Range("B10").Interior.color = RGB(0, 255, 0) ' Verde
        End If
        
        .Range("A11").Value = "Cpk [ Capacidad real del proceso ]"
        .Range("A11").Font.Bold = True
        .Range("B11").Value = Cpk
        .Range("B11").NumberFormat = "0.0000"
        ' Aplica la alineación horizontal: Centro
        .Range("B11").HorizontalAlignment = xlCenter
        ' Aplica la alineación vertical (opcional, pero buena práctica)
        .Range("B11").VerticalAlignment = xlCenter
        
        If Cpk < 1.33 Then
            .Range("B11").Interior.color = RGB(255, 0, 0)
        ElseIf Cpk < 1.67 Then
            .Range("B11").Interior.color = RGB(255, 255, 0)
        Else
            .Range("B11").Interior.color = RGB(0, 255, 0)
        End If
        
        .Range("A12").Value = "Pp [ Capacidad potencial basada en datos reales ]"
        .Range("A12").Font.Bold = True
        .Range("B12").Value = Pp
        .Range("B12").NumberFormat = "0.0000"
        ' Aplica la alineación horizontal: Centro
        .Range("B12").HorizontalAlignment = xlCenter
        ' Aplica la alineación vertical (opcional, pero buena práctica)
        .Range("B12").VerticalAlignment = xlCenter
        
        
        If Pp < 1.33 Then
            .Range("B12").Interior.color = RGB(255, 0, 0)
        ElseIf Pp < 1.67 Then
            .Range("B12").Interior.color = RGB(255, 255, 0)
        Else
            .Range("B12").Interior.color = RGB(0, 255, 0)
        End If
        
        .Range("A13").Value = "Ppk [ Capacidad real basada en datos reales ]"
        .Range("A13").Font.Bold = True
        .Range("B13").Value = Ppk
        .Range("B13").NumberFormat = "0.0000"
        ' Aplica la alineación horizontal: Centro
        .Range("B13").HorizontalAlignment = xlCenter
        ' Aplica la alineación vertical (opcional, pero buena práctica)
        .Range("B13").VerticalAlignment = xlCenter
        
        If Ppk < 1.33 Then
            .Range("B13").Interior.color = RGB(255, 0, 0)
        ElseIf Ppk < 1.67 Then
            .Range("B13").Interior.color = RGB(255, 255, 0)
        Else
            .Range("B13").Interior.color = RGB(0, 255, 0)
        End If
        
        If Target <> 0 Then
            .Range("A14").Value = "Cpm [ Capacidad centrada en el objetivo (Target) ]"
            .Range("A14").Font.Bold = True
            .Range("B14").Value = Cpm
            .Range("B14").NumberFormat = "0.0000"
            
            ' Aplica la alineación horizontal: Centro
            .Range("B14").HorizontalAlignment = xlCenter
            ' Aplica la alineación vertical (opcional, pero buena práctica)
            .Range("B14").VerticalAlignment = xlCenter
            
            
            If Cpm < 1.33 Then
                .Range("B14").Interior.color = RGB(255, 0, 0)
            ElseIf Cpm < 1.67 Then
                .Range("B14").Interior.color = RGB(255, 255, 0)
            Else
                .Range("B14").Interior.color = RGB(0, 255, 0)
            End If
        End If
        
        ' Sección 3: Interpretación
        .Range("A16").Value = "Interpretación"
        .Range("A16").Font.Bold = True
        .Range("A16").Font.Size = 12
        .Range("A16:B16").Merge
        .Range("A16:B16").Interior.color = RGB(200, 200, 200)
        
        .Range("A17").Value = "Prueba de Normalidad:"
        .Range("A17").Font.Bold = True
        .Range("B17").Value = IIf(IsNormal, "Datos normales", "ADVERTENCIA: Datos no normales")
        
        If Not IsNormal Then
            .Range("B17").Font.color = RGB(255, 0, 0)
            .Range("B17").Font.Bold = True
            ' Aplica la alineación horizontal: Centro
            .Range("B17").HorizontalAlignment = xlCenter
            ' Aplica la alineación vertical (opcional, pero buena práctica)
            .Range("B17").VerticalAlignment = xlCenter
        End If
        
        .Range("A18").Value = "Recomendación:"
        .Range("A18").Font.Bold = True
        .Range("B18").Value = IIf(IsNormal, "Índices Cp/Cpk válidos", "Usar Pp/Ppk o métodos no paramétricos")
        
        If Not IsNormal Then
            .Range("B18").Font.color = RGB(255, 0, 0)
            .Range("B18").Font.Bold = True
            ' Aplica la alineación horizontal: Centro
            .Range("B18").HorizontalAlignment = xlCenter
            ' Aplica la alineación vertical (opcional, pero buena práctica)
            .Range("B18").VerticalAlignment = xlCenter
        End If
        
        ' Formatear bordes
        .Range("A3:B7").Borders.LineStyle = xlContinuous
        .Range("A9:B13").Borders.LineStyle = xlContinuous
        If Target <> 0 Then
            .Range("A14:B14").Borders.LineStyle = xlContinuous
        End If
        .Range("A16:B18").Borders.LineStyle = xlContinuous
        
        ' Ajustar ancho de columnas al contenido
        ws.Columns("A:B").AutoFit
    End With
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
End Sub
