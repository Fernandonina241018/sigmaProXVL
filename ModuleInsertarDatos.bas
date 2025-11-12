'================================================================================
' MÓDULO: Análisis Estadístico Profesional - Layout Optimizado
' VERSIÓN: 3.0 - Diseño Horizontal con Separación de Gráficos
' DESCRIPCIÓN: Informe tipo revista con columnas izquierda protegidas (A-G)
' MEJORAS V3.0:
'   - Resultados inician en columna H (protege columnas A-G)
'   - Layout horizontal que evita superposición de gráficos
'   - Diseño tipo revista profesional con secciones verticales
'   - Gráficos posicionados estratégicamente debajo de cada columna
'================================================================================
''================================================================================
Public Sub MostrarAnalisisVerticalEnHoja()
    '----------------------------------------------------------------------------
    ' SECCIÓN 1: DECLARACIÓN DE VARIABLES
    '----------------------------------------------------------------------------
    Dim wb As Workbook
    Dim rango As Range
    Dim ws As Worksheet
    Dim fila As Long, colResultados As Integer, col As Long
    Dim numColumnas As Integer
    Dim stats() As EstadisticasColumna
    Dim nombreHoja As String
    Dim numeroHoja As Integer
    Dim exp As Double
    Dim mediana As Double
    Dim varianza As Double
    exp = 4#
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

    '----------------------------------------------------------------------------
    ' SECCIÓN 2: CONFIGURACIÓN INICIAL Y VALIDACIONES
    '----------------------------------------------------------------------------
    ' Referenciar libro activo
    Set wb = ActiveWorkbook

    ' Obtener rango desde control RefEdit (interfaz de usuario)
    On Error Resume Next
    Set rango = Range(sigmaproxvl.txtRango.Text)
    On Error GoTo 0

    ' VALIDACIÓN 1: Rango válido
    If rango Is Nothing Then
        Debug.Print "Rango inválido. Por favor seleccione un rango válido.", vbExclamation
        Exit Sub
    End If

    ' Determinar workbook correcto (maneja rangos entre workbooks)
    Set wb = rango.Parent.Parent

    ' VALIDACIÓN 2: Número de columnas
    numColumnas = rango.Columns.count
    If numColumnas = 0 Then
        Debug.Print "No se encontraron columnas en el rango seleccionado.", vbExclamation
        Exit Sub
    End If

    '----------------------------------------------------------------------------
    ' SECCIÓN 3: ANÁLISIS ESTADÍSTICO POR COLUMNA
    '----------------------------------------------------------------------------
    ReDim stats(1 To numColumnas)
    For col = 1 To numColumnas
        stats(col) = AnalizarColumna(rango.Columns(col))
    Next col

    '----------------------------------------------------------------------------
    ' SECCIÓN 4: CREACIÓN Y CONFIGURACIÓN DE HOJA DE RESULTADOS
    '----------------------------------------------------------------------------
    ' Generar nombre único para hoja de resultados
    numeroHoja = ObtenerProximoNumeroHoja(wb, "Análisis Estadístico")
    nombreHoja = "Análisis Estadístico " & numeroHoja

    ' Crear nueva hoja al final del workbook
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.count))
    ws.Name = nombreHoja

    Dim minimoGlobal As Double
    Dim maximoGlobal As Double

    ' Inicializar con valores del primer columna válida
    Dim primeraColumnaValida As Integer
    For col = 1 To numColumnas
        If stats(col).count > 0 Then
            minimoGlobal = stats(col).minimo
            maximoGlobal = stats(col).maximo
            primeraColumnaValida = col
            Exit For
        End If
    Next col

    ' Buscar mínimo y máximo globales
    For col = primeraColumnaValida + 1 To numColumnas
        If stats(col).count > 0 Then
            If stats(col).minimo < minimoGlobal Then minimoGlobal = stats(col).minimo
            If stats(col).maximo > maximoGlobal Then maximoGlobal = stats(col).maximo
        End If
    Next col
    '----------------------------------------------------------------------------
    ' SECCIÓN 5: FORMATO DE REPORTE - ENCABEZADO
    '----------------------------------------------------------------------------
    With ws
        '-----------------------------------------------------------------------
        ' CONFIGURACIÓN DE PÁGINA PARA IMPRESIÓN
        '-----------------------------------------------------------------------
        With .PageSetup
            .Orientation = xlPortrait
            .PaperSize = xlPaperLetter
            .LeftMargin = Application.InchesToPoints(0.75)
            .RightMargin = Application.InchesToPoints(0.75)
            .TopMargin = Application.InchesToPoints(1)
            .BottomMargin = Application.InchesToPoints(1)
            .HeaderMargin = Application.InchesToPoints(0.5)
            .FooterMargin = Application.InchesToPoints(0.5)
            .PrintGridlines = False
            .PrintHeadings = False
            .CenterHorizontally = True
        End With
        
        '-----------------------------------------------------------------------
        ' MEMBRETE CORPORATIVO (Filas 1-3)
        '-----------------------------------------------------------------------
        ' Espacio para logo o nombre de empresa
        .Range("A1:E1").Merge
        With .Range("A1")
            .Value = "[LABORATORIO SUED S.R.L. / VALIDACIONES]"
            .Font.Name = "Segoe UI"
            .Font.Size = 16
            .Font.Bold = True
            .Font.color = RGB(31, 78, 120)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Rows(1).RowHeight = 25
        
        ' Subtítulo del documento
        .Range("A2:E2").Merge
        With .Range("A2")
            .Value = "REPORTE DE ANÁLISIS ESTADÍSTICO"
            .Font.Name = "Segoe UI"
            .Font.Size = 12
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Rows(2).RowHeight = 20
        
        ' Línea separadora decorativa
        With .Range("A3:E3")
            .Merge
            .Borders(xlEdgeBottom).LineStyle = xlDouble
            .Borders(xlEdgeBottom).Weight = xlThick
            .Borders(xlEdgeBottom).ColorIndex = 1
        End With
        .Rows(3).RowHeight = 5
        
        '-----------------------------------------------------------------------
        ' BLOQUE 1: IDENTIFICACIÓN DEL DOCUMENTO (Filas 4-5)
        '-----------------------------------------------------------------------
        .Rows(4).RowHeight = 8  ' Espacio en blanco
        
        Call FormatearSeccionTitulo(.Range("A5:E5"), "IDENTIFICACIÓN DEL ANÁLISIS")
        
        ' Número de análisis
        Call FormatearEtiqueta(.Range("A6:B6"))
        .Range("A6").Value = "Número de Análisis:"
        Call FormatearDato(.Range("C6:E6"))
        .Range("C6").Value = "# " & Format(numeroHoja, "0000")
        .Range("C6").VerticalAlignment = xlCenter
        
        ' Hoja de trabajo
        Call FormatearEtiqueta(.Range("A7:B7"))
        .Range("A7").Value = "Hoja de Trabajo:"
        Call FormatearDato(.Range("C7:E7"))
        .Range("C7").Value = ws.Name
        
        ' Fecha y hora del análisis
        Call FormatearEtiqueta(.Range("A8:B8"))
        .Range("A8").Value = "Fecha y Hora:"
        Call FormatearDato(.Range("C8:E8"))
        .Range("C8").Value = Format(Now, "dddd, dd 'de' mmmm 'de' yyyy") & _
                            vbLf & Format(Now, "hh:mm:ss AM/PM")
        .Rows(8).RowHeight = 30
        
        '-----------------------------------------------------------------------
        ' BLOQUE 2: INFORMACIÓN DEL DATASET (Filas 9-12)
        '-----------------------------------------------------------------------
        .Rows(9).RowHeight = 8  ' Espacio en blanco
        
        Call FormatearSeccionTitulo(.Range("A10:E10"), "INFORMACIÓN DEL CONJUNTO DE DATOS")
        
        ' Rango analizado
        Call FormatearEtiqueta(.Range("A11:B11"))
        .Range("A11").Value = "Rango Analizado:"
        Call FormatearDato(.Range("C11:E11"))
        .Range("C11").Value = rango.Address(False, False)
        
        ' Dimensiones
        Call FormatearEtiqueta(.Range("A12:B12"))
        .Range("A12").Value = "Columnas Analizadas:"
        Call FormatearDato(.Range("C12:E12"))
        .Range("C12").Value = numColumnas & " columna(s)"
        
        ' Total de datos procesados
        Call FormatearEtiqueta(.Range("A13:B13"))
        .Range("A13").Value = "Total de Registros:"
        Call FormatearDato(.Range("C13:E13"))
        totalDatos = 0
        For col = 1 To numColumnas
            totalDatos = totalDatos + stats(col).count
        Next col
        .Range("C13").Value = Format(totalDatos, "#,##0") & " registro(s)"
        
        '-----------------------------------------------------------------------
        ' BLOQUE 3: ESTADÍSTICAS GENERALES (Filas 14-17)
        '-----------------------------------------------------------------------
        .Rows(14).RowHeight = 8  ' Espacio en blanco
        
        Call FormatearSeccionTitulo(.Range("A15:E15"), "ESTADÍSTICAS GENERALES")
        
        ' Máximo Global
        Call FormatearEtiqueta(.Range("A16:B16"))
        .Range("A16").Value = "Valor Máximo Global:"
        Call FormatearDato(.Range("C16:E16"))
        .Range("C16").Value = Format(maximoGlobal, "#,##0.00")
        .Range("C16").Font.Bold = True
        
        ' Mínimo Global
        Call FormatearEtiqueta(.Range("A17:B17"))
        .Range("A17").Value = "Valor Mínimo Global:"
        Call FormatearDato(.Range("C17:E17"))
        .Range("C17").Value = Format(minimoGlobal, "#,##0.00")
        .Range("C17").Font.Bold = True
        
        ' Rango de valores
        Call FormatearEtiqueta(.Range("A18:B18"))
        .Range("A18").Value = "Amplitud de Rango:"
        Call FormatearDato(.Range("C18:E18"))
        .Range("C18").Value = Format(maximoGlobal - minimoGlobal, "#,##0.00")
        
        '-----------------------------------------------------------------------
        ' BLOQUE 4: DETECCIÓN DE ANOMALÍAS (Filas 19-22)
        '-----------------------------------------------------------------------
        .Rows(19).RowHeight = 8  ' Espacio en blanco
        
        Call FormatearSeccionTitulo(.Range("A20:E20"), "DETECCIÓN DE ANOMALÍAS (OUTLIERS)")
        
        ' Calcular total de outliers
        totalOutliers = 0
        For col = 1 To numColumnas
            totalOutliers = totalOutliers + stats(col).NumOutliers
        Next col
        
        ' Outliers detectados
        Call FormatearEtiqueta(.Range("A21:B21"))
        .Range("A21").Value = "Outliers Detectados:"
        Call FormatearDato(.Range("C21:E21"))
        .Range("C21").Value = Format(totalOutliers, "#,##0") & " registro(s)"
        If totalOutliers > 0 Then
            .Range("C21").Font.color = RGB(192, 0, 0)  ' Rojo
            .Range("C21").Font.Bold = True
        End If
        
        ' Porcentaje de outliers
        Call FormatearEtiqueta(.Range("A22:B22"))
        .Range("A22").Value = "Porcentaje de Anomalías:"
        Call FormatearDato(.Range("C22:E22"))
        If totalDatos > 0 Then
            porcentajeOutliers = (totalOutliers / totalDatos) * 100
            .Range("C22").Value = Format(porcentajeOutliers, "0.00") & " %"
            If porcentajeOutliers > 5 Then
                .Range("C22").Font.color = RGB(192, 0, 0)  ' Rojo
                .Range("C22").Font.Bold = True
            End If
        Else
            .Range("C22").Value = "N/A"
        End If
        
        '-----------------------------------------------------------------------
        ' BLOQUE 5: EVALUACIÓN Y CONCLUSIONES (Filas 23-26)
        '-----------------------------------------------------------------------
        .Rows(23).RowHeight = 8  ' Espacio en blanco
        
        Call FormatearSeccionTitulo(.Range("A24:E24"), "EVALUACIÓN DE CALIDAD DE DATOS")
        
        ' Estado de normalización
        Call FormatearEtiqueta(.Range("A25:B25"))
        .Range("A25").Value = "Estado de Normalización:"
        Call FormatearDato(.Range("C25:E25"))
        If totalOutliers = 0 Then
            .Range("C25").Value = "? DATOS NORMALIZADOS"
            .Range("C25").Font.color = RGB(0, 128, 0)  ' Verde
            .Range("C25").Font.Bold = True
            .Range("C25").Interior.color = RGB(198, 239, 206)  ' Fondo verde claro
        Else
            .Range("C25").Value = "? DATOS NO NORMALIZADOS"
            .Range("C25").Font.color = RGB(192, 0, 0)  ' Rojo
            .Range("C25").Font.Bold = True
            .Range("C25").Interior.color = RGB(255, 199, 206)  ' Fondo rojo claro
        End If
        
        ' Recomendación
        Call FormatearEtiqueta(.Range("A26:B26"))
        .Range("A26").Value = "Recomendación:"
        Call FormatearDato(.Range("C26:E26"))
        If totalOutliers = 0 Then
            .Range("C26").Value = "Los datos cumplen con los estándares de calidad." & _
                                vbLf & "No se requiere acción correctiva."
        Else
            .Range("C26").Value = "Se recomienda revisar los " & totalOutliers & _
                                " registro(s) identificados como outliers." & _
                                vbLf & "Evaluar si son errores o valores legítimos."
        End If
        .Rows(26).RowHeight = 30
        
        '-----------------------------------------------------------------------
        ' PIE DE PÁGINA DEL REPORTE (Fila 27)
        '-----------------------------------------------------------------------
        .Rows(27).RowHeight = 8  ' Espacio en blanco
        
        With .Range("A28:E28")
            .Merge
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeTop).ColorIndex = 1
            .Value = "Fin del Encabezado del Reporte - Datos Detallados a Continuación"
            .Font.Name = "Arial"
            .Font.Size = 8
            .Font.Italic = True
            .Font.color = RGB(128, 128, 128)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Rows(28).RowHeight = 15
        
        .Rows(29).RowHeight = 12  ' Espacio final antes de datos
        
        '-----------------------------------------------------------------------
        ' AJUSTES FINALES DE FORMATO PARA IMPRESIÓN
        '-----------------------------------------------------------------------
        ' Configurar ancho de columnas (respetando diseño A:E)
        .Columns("A:B").ColumnWidth = 20
        .Columns("C:E").ColumnWidth = 25
        
        ' Establecer área de impresión solo para el encabezado
        .PageSetup.PrintArea = "$A$1:$E$28"
        
        ' Configurar encabezado y pie de página de impresión
        With .PageSetup
            .LeftHeader = ""
            .CenterHeader = "&""Arial,Bold""&12" & ws.Name
            .RightHeader = ""
            .LeftFooter = "&""Arial""&8Generado: " & Format(Now, "dd/mm/yyyy hh:mm")
            .CenterFooter = ""
            .RightFooter = "&""Arial""&8Página &P de &N"
        End With
        
        ' Borde exterior completo para el reporte
        With .Range("A1:E28")
            .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, ColorIndex:=1
        End With
        '----------------------------------------------------------------------------
        ' SECCIÓN 6: PRESENTACIÓN DE RESULTADOS POR COLUMNA
        '----------------------------------------------------------------------------
        colResultados = 7 ' Columna inicial para resultados (C)

        For col = 1 To numColumnas
            '------------------------------------------------------------------------
            ' SUBSECCIÓN 6.1: ENCABEZADO DE COLUMNA
            '------------------------------------------------------------------------
            fila = 1 ' Fila inicial para cada columna

            ' Título con información de límites del formulario
            .Cells(fila, colResultados).Value = stats(col).NombreColumna & " ( Rango " & stats(col).columna & ")"
            Range(.Cells(fila, colResultados), .Cells(fila, colResultados + 3)).Merge
            Call FormatearEncabezado(.Cells(fila, colResultados))
            Call FormatearEncabezado(.Cells(fila, colResultados + 1))


            '------------------------------------------------------------------------
            ' SUBSECCIÓN 6.2: ESTADÍSTICOS BÁSICOS
            '------------------------------------------------------------------------
            If stats(col).count > 0 Then
                Dim rngDatos As Range


                 ' Límites en UserForm
                fila = fila + 1
                .Cells(fila, colResultados).Value = "Estadisticos: "
                .Cells(fila, colResultados + 1).Value = "Datos [Resultados]"
                .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                Call FormatearEncabezado(.Cells(fila, colResultados))
                Call FormatearEncabezado(.Cells(fila, colResultados + 1))

                ' Límites en UserForm
                fila = fila + 1
                .Cells(fila, colResultados).Value = "Límites: "
                .Cells(fila, colResultados + 1).Value = sigmaproxvl.cboLimiteSuperior & " - " & sigmaproxvl.cboLimiteInferior
                .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                Call FormatearEncabezado(.Cells(fila, colResultados))
                Call FormatearDato(.Cells(fila, colResultados + 1))

                ' Count (n)
                fila = fila + 1
                .Cells(fila, colResultados).Value = "Count Date(n):"
                .Cells(fila, colResultados + 1).Value = stats(col).count
                .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                Call FormatearEncabezado(.Cells(fila, colResultados))
                Call FormatearDato(.Cells(fila, colResultados + 1))

                ' Promedio
                fila = fila + 1
                .Cells(fila, colResultados).Value = "Promedio:"
                .Cells(fila, colResultados + 1).Value = stats(col).promedio
                .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                Call FormatearEncabezado(.Cells(fila, colResultados))
                Call FormatearDato(.Cells(fila, colResultados + 1))

                ' Desviación Estándar con coloración condicional
                fila = fila + 1
                .Cells(fila, colResultados).Value = "Desv. Estándar:"
                .Cells(fila, colResultados + 1).Value = stats(col).desviacionEstandar
                .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                Call FormatearEncabezado(.Cells(fila, colResultados))
                Call FormatearDato(.Cells(fila, colResultados + 1))

                ' COLORACIÓN: Rojo si > 4, Verde si <= 4
                If stats(col).desviacionEstandar > exp Then
                    .Cells(fila, colResultados + 1).Interior.color = RGB(255, 0, 0)
                Else
                    .Cells(fila, colResultados + 1).Interior.color = RGB(0, 255, 0)
                End If

                ' CONFIRMACION PARA MEDIANA
                fila = fila + 1
                .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                Call FormatearEncabezado(.Cells(fila, colResultados))
                Call FormatearDato(.Cells(fila, colResultados + 1))

                If stats(col).count Mod 2 = 0 Then
                '.Cells(fila, colResultados + 1).Value = stats(col).count
                    .Cells(fila, colResultados).Value = "Mediana --> [Par]"
                    .Cells(fila, colResultados + 1).Value = stats(col).mediana
                Else
                    .Cells(fila, colResultados).Value = "Mediana --> [Impar]"
                    .Cells(fila, colResultados + 1).Value = stats(col).mediana
                End If

                ' %RSD
                fila = fila + 1
                .Cells(fila, colResultados).Value = "RSD (%):"
                .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                Call FormatearEncabezado(.Cells(fila, colResultados))
                Call FormatearDato(.Cells(fila, colResultados + 1))
                .Cells(fila, colResultados + 1).Value = (stats(col).desviacionEstandar / Abs(stats(col).promedio)) * 100

                ' Máximo con coloración condicional
                fila = fila + 1
                .Cells(fila, colResultados).Value = "Máximo:"
                .Cells(fila, colResultados + 1).Value = stats(col).maximo
                .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                Call FormatearEncabezado(.Cells(fila, colResultados))
                Call FormatearDato(.Cells(fila, colResultados + 1))
                If stats(col).maximo > sigmaproxvl.cboLimiteSuperior Then
                    .Cells(fila, colResultados + 1).Interior.color = RGB(255, 0, 0)
                Else
                    .Cells(fila, colResultados + 1).Interior.color = RGB(0, 255, 0)
                End If

                ' Mínimo con coloración condicional
                fila = fila + 1
                .Cells(fila, colResultados).Value = "Mínimo:"
                .Cells(fila, colResultados + 1).Value = stats(col).minimo
                .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                Call FormatearEncabezado(.Cells(fila, colResultados))
                Call FormatearDato(.Cells(fila, colResultados + 1))
                If stats(col).minimo < sigmaproxvl.cboLimiteInferior Then
                    .Cells(fila, colResultados + 1).Interior.color = RGB(255, 0, 0)   ' Rojo
                Else
                    .Cells(fila, colResultados + 1).Interior.color = RGB(0, 255, 0)   ' Verde
                End If

                ' Varianza
                fila = fila + 1
                .Cells(fila, colResultados).Value = "Varianza:"
                .Cells(fila, colResultados + 1).Value = stats(col).varianza
                .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                Call FormatearEncabezado(.Cells(fila, colResultados))
                Call FormatearDato(.Cells(fila, colResultados + 1))

                ' Moda
                fila = fila + 1
                .Cells(fila, colResultados).Value = "Moda:"
                .Cells(fila, colResultados + 1).Value = stats(col).moda
                .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                Call FormatearEncabezado(.Cells(fila, colResultados))
                Call FormatearDato(.Cells(fila, colResultados + 1))

                '------------------------------------------------------------------------
                ' SUBSECCIÓN 6.5: CÁLCULOS ESPECIALES - VALIDACIÓN FARMACÉUTICA
                '------------------------------------------------------------------------
                ' Cálculo de Expectativa Matemática
                Dim expmath As Double
                On Error Resume Next
                expmath = CDbl(sigmaproxvl.cboExpectativa.Value)
                On Error GoTo 0

                If expmath <> 0 Then
                    fila = fila + 1
                    .Cells(fila, colResultados).Value = "Expectativa Math:"
                    .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                    Call FormatearEncabezado(.Cells(fila, colResultados))
                    Call FormatearDato(.Cells(fila, colResultados + 1))

                    ' Cálculo de porcentaje de cumplimiento
                    If expmath <= stats(col).promedio Then
                        .Cells(fila, colResultados + 1).Value = _
                            ((expmath / Abs(stats(col).promedio)) * 100) & "%"
                    Else
                        .Cells(fila, colResultados + 1).Value = _
                            (Abs((stats(col).promedio) / expmath) * 100) & "%"
                    End If

                    ' Coloración: Rojo si < 95%, Verde si >= 95%
                    If expmath < 0.95 Then
                        .Cells(fila, colResultados + 1).Interior.color = RGB(255, 0, 0)
                    Else
                        .Cells(fila, colResultados + 1).Interior.color = RGB(0, 255, 0)
                    End If
                Else
                    fila = fila + 1
                    .Cells(fila, colResultados).Value = "Expectativa Math:"
                    .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                    Call FormatearEncabezado(.Cells(fila, colResultados))
                    Call FormatearDato(.Cells(fila, colResultados + 1))
                    .Cells(fila, colResultados + 1).Value = "No Aplica"
                End If

                ' Asimetría
                fila = fila + 1
                .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                Call FormatearEncabezado(.Cells(fila, colResultados))
                Call FormatearDato(.Cells(fila, colResultados + 1))
                If stats(col).asimetria < 0 Then
                    .Cells(fila, colResultados).Value = "Asimetría (Skewness) --> Izquierda"
                    .Cells(fila, colResultados + 1).Value = stats(col).asimetria
                ElseIf stats(col).asimetria = 0 Then
                    .Cells(fila, colResultados).Value = "Asimetría (Skewness) --> Centro"
                    .Cells(fila, colResultados + 1).Value = stats(col).asimetria
                Else:
                    .Cells(fila, colResultados).Value = "Asimetría (Skewness) --> Derecha"
                    .Cells(fila, colResultados + 1).Value = stats(col).asimetria
                End If

                ' Curtosis
                fila = fila + 1
                .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                Call FormatearEncabezado(.Cells(fila, colResultados))
                Call FormatearDato(.Cells(fila, colResultados + 1))
                If stats(col).asimetria < 0 Then
                    .Cells(fila, colResultados).Value = "Curtosis (Kurtosis) --> Platicúrtica (Pico Plano)"
                    .Cells(fila, colResultados + 1).Value = stats(col).curtosis
                ElseIf stats(col).asimetria = 0 Then
                    .Cells(fila, colResultados).Value = "Curtosis (Kurtosis) --> Distribución Norma"
                    .Cells(fila, colResultados + 1).Value = stats(col).curtosis
                Else:
                    .Cells(fila, colResultados).Value = "Curtosis (Kurtosis) --> Leptocúrtica (Pico Agudo)"
                    .Cells(fila, colResultados + 1).Value = stats(col).curtosis
                End If

                '------------------------------------------------------------------------
                ' SUBSECCIÓN 6.6: CÁLCULO DE F0 (STERILIZATION VALUE)
                ' PROPÓSITO: Validación de procesos de esterilización
                ' REFERENCIA: USP <1229.5> - Sterilization Validation
                '------------------------------------------------------------------------

                Dim std As Double, fn As Double
                Dim f0_raw As Double
                Dim f0 As Double
                Dim z As Integer, t0 As Integer

                ' Parámetros estándar para cálculo F0
                std = 121   ' Temperatura de referencia (°C)
                z = 10      ' Valor z (°C)
                t0 = 10     ' Tiempo de referencia (minutos)

                ' FÓRMULA F0: ?10^((T-121)/z) dt
                On Error Resume Next

                fn = t0 ^ ((stats(col).promedio - std) / z)

                f0_raw = ((CDate(sigmaproxvl.cboTiempoFinal.Value) - CDate(sigmaproxvl.cboTiempoInicio.Value)) * 1440) * _
                    t0 ^ ((stats(col).promedio - std) / z)

                On Error GoTo 0

                f0 = Abs(f0_raw) ' Valor absoluto para F0

                ' Presentación de resultado F0
                fila = fila + 1
                .Cells(fila, colResultados).Value = "[ F0: ] ---> "
                If sigmaproxvl.cboModoAnalisis.Value = "Esterilización" Then
                    .Cells(fila, colResultados + 1).Value = " [ " & fn & " / " & f0 & " ] "
                Else
                    .Cells(fila, colResultados + 1).Value = "No Aplica"
                End If
                .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                Call FormatearEncabezado(.Cells(fila, colResultados))
                Call FormatearDato(.Cells(fila, colResultados + 1))

                ' Coloración: Rojo si F0 < 15, Verde si F0 >= 15
                If f0 < 15 Then
                    .Cells(fila, colResultados + 1).Interior.color = RGB(255, 0, 0)
                Else
                    .Cells(fila, colResultados + 1).Interior.color = RGB(0, 255, 0)
                End If

            Else
                '------------------------------------------------------------------------
                ' SUBSECCIÓN 6.7: CASO SIN DATOS VÁLIDOS
                '------------------------------------------------------------------------
                fila = fila + 1
                .Cells(fila, colResultados).Value = "Sin datos numéricos"
                .Range(.Cells(fila, colResultados), .Cells(fila, colResultados + 1)).Merge
                fila = fila + 5 ' Espacio adicional para mantener formato
            End If

                '------------------------------------------------------------------------
                ' SUBSECCIÓN 6.3: DETECCIÓN Y REPORTE DE OUTLIERS
                '------------------------------------------------------------------------
                If stats(col).NumOutliers > 0 Then
                    ' Encabezado de sección outliers
                    fila = fila + 1
                    .Cells(fila, colResultados).Value = "OUTLIERS DETECTADOS:"
                    .Range(.Cells(fila, colResultados), .Cells(fila, colResultados + 3)).Merge
                    Call FormatearEncabezado(.Cells(fila, colResultados))
                    Call FormatearEncabezado(.Cells(fila, colResultados + 1))

                    ' Cantidad de outliers
                    fila = fila + 1
                    .Cells(fila, colResultados).Value = "Cantidad:"
                    .Cells(fila, colResultados + 1).Value = stats(col).NumOutliers & " de " & stats(col).count
                    .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                    Call FormatearEncabezado(.Cells(fila, colResultados))
                    Call FormatearDato(.Cells(fila, colResultados + 1))
                    Call AjustarAlturaPorContenido(.Cells(fila, colResultados + 1))

                    ' Valores específicos de outliers
                    fila = fila + 1
                    .Cells(fila, colResultados).Value = "Valores:"
                    .Cells(fila, colResultados + 1).Value = ArrayToString(stats(col).Outliers)
                    .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                    Call FormatearEncabezado(.Cells(fila, colResultados))
                    Call FormatearDato(.Cells(fila, colResultados + 1))
                    .Cells(fila, colResultados + 1).WrapText = True
                    .Cells(fila, colResultados + 1).Font.Size = 8 ' Fuente más pequeña para valores
                    Call AjustarAlturaPorTextoCombinado(.Cells(fila, colResultados + 1))


                    ' Límites IQR utilizados
                    fila = fila + 1
                    .Cells(fila, colResultados).Value = "Límites IQR:"
                    .Cells(fila, colResultados + 1).Value = "[" & _
                        Format(stats(col).LimiteInferiorOutlier, "0.0000") & " - " & _
                        Format(stats(col).LimiteSuperiorOutlier, "0.0000") & "]"
                    .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                    Call FormatearEncabezado(.Cells(fila, colResultados))
                    Call FormatearDato(.Cells(fila, colResultados + 1))

                    '------------------------------------------------------------------------
                    ' SUBSECCIÓN 6.4: ESTADÍSTICAS ROBUSTAS (EXCLUYENDO OUTLIERS)
                    '------------------------------------------------------------------------
                    fila = fila + 1
                    .Cells(fila, colResultados).Value = "--- ESTADÍSTICAS ROBUSTAS ---"
                    .Range(.Cells(fila, colResultados), .Cells(fila, colResultados + 3)).Merge
                    Call FormatearEncabezado(.Cells(fila, colResultados))
                    Call FormatearEncabezado(.Cells(fila, colResultados + 1))

                    ' Media Robusta
                    fila = fila + 1
                    .Cells(fila, colResultados).Value = "Media robusta:"
                    .Cells(fila, colResultados + 1).Value = stats(col).MediaRobusta
                    .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                    Call FormatearEncabezado(.Cells(fila, colResultados))
                    Call FormatearDato(.Cells(fila, colResultados + 1))
                    .Cells(fila, colResultados + 1).Interior.color = RGB(173, 216, 230) ' Azul claro

                    ' Desviación Estándar Robusta
                    fila = fila + 1
                    .Cells(fila, colResultados).Value = "DE robusta:"
                    .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                    Call FormatearEncabezado(.Cells(fila, colResultados))
                    Call FormatearDato(.Cells(fila, colResultados + 1))
                    .Cells(fila, colResultados + 1).Value = stats(col).DesvEstandarRobusta
                    .Cells(fila, colResultados + 1).Interior.color = RGB(173, 216, 230)

                    ' %RSD Robusto
                    fila = fila + 1
                    .Cells(fila, colResultados).Value = "RSD robusto (%):"
                    .Range(.Cells(fila, colResultados + 1), .Cells(fila, colResultados + 3)).Merge
                    Call FormatearEncabezado(.Cells(fila, colResultados))
                    Call FormatearDato(.Cells(fila, colResultados + 1))
                    .Cells(fila, colResultados + 1).Value = stats(col).RSDrobusto & "%"
                    .Cells(fila, colResultados + 1).Interior.color = RGB(173, 216, 230)
                Else
                    ' Mensaje: Sin outliers detectados
                    fila = fila + 1
                    .Cells(fila, colResultados).Value = "? Sin outliers detectados"
                    .Range(.Cells(fila, colResultados), .Cells(fila, colResultados + 3)).Merge
                    Call FormatearEncabezado(.Cells(fila, colResultados))
                    Call FormatearEncabezado(.Cells(fila, colResultados + 1))
                   fila = fila + 1
                End If

            '------------------------------------------------------------------------
            ' SUBSECCIÓN 6.8: FORMATEO FINAL DEL BLOQUE DE COLUMNA
            '------------------------------------------------------------------------
            ' Aplicar bordes al bloque completo de resultados
            .Range(.Cells(1, colResultados), .Cells(fila, colResultados + 3)).Borders.LineStyle = xlContinuous

            ' Formato numérico para celdas de valores
            If stats(col).count > 0 Then
                '.Range(.Cells(8, colResultados + 1), .Cells(fila, colResultados + 1)).NumberFormat = "0.0000"
            End If

            ' Ajuste de anchos de columna para mejor visualización
            .Columns(colResultados).ColumnWidth = 20
            .Columns(colResultados + 1).ColumnWidth = 12
            .Columns(colResultados + 2).ColumnWidth = 8

            '------------------------------------------------------------------------
            ' SUBSECCIÓN 6.9: LLAMADA A MÓDULO DE GRÁFICOS
            '------------------------------------------------------------------------
            If stats(col).count > 1 Then ' Requisito mínimo: 2 puntos para gráficos
                Dim graficoTop As Long
                graficoTop = fila + 2 ' Espacio después de la tabla

                ' Llamada al módulo externo de gráficos
                Call CrearGraficosParaColumna(ws, stats(col), graficoTop, colResultados)
            End If

            ' Avanzar a la siguiente posición de columna
            colResultados = colResultados + 8 ' Espacio entre columnas de resultados
        Next col
        '----------------------------------------------------------------------------
        ' SECCIÓN 7: RESUMEN GENERAL DEL ANÁLISIS
        '--------------------------------------------------------

        '----------------------------------------------------------------------------
        ' SECCIÓN 8: AJUSTES FINALES DE FORMATEO
        '----------------------------------------------------------------------------
        .Columns("A:ZZ").AutoFit ' Autoajustar todas las columnas utilizadas
        '.Rows("1:100").AutoFit
    End With

    '----------------------------------------------------------------------------
    ' SECCIÓN 9: ACTIVACIÓN Y PREPARACIÓN FINAL
    '----------------------------------------------------------------------------
    ws.Activate           ' Activar la hoja de resultados
    ws.Range("A1").Select ' Posicionar cursor en celda A1
    ActiveWindow.DisplayGridlines = False ' Se aplica a la ventana activa

    Call AjustarColumnasHojaActiva

    '----------------------------------------------------------------------------
    ' SECCIÓN 10: ANÁLISIS DE CORRELACIÓN (CONDICIONAL)
    '----------------------------------------------------------------------------
    If sigmaproxvl.chkCorrelacion.Value Then
        Call EjecutarAnalisisCorrelacion(stats, wb, True)
    End If

    ' Después del análisis principal
    If sigmaproxvl.chkCapacidadProceso.Value Then
        Call RunCapabilityAnalysis
    End If

    ' NOTA: Mensaje desactivado para procesos automatizados
    ' debug.print "Análisis completado correctamente. Outliers detectados: " & totalOutliers, vbInformation
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
End Sub

' VERSIÓN 2: DESVIACIÓN PORCENTUAL REAL
' Cálculo matemáticamente correcto y simétrico
' ============================================

Public Sub AplicarColorDesviacion(celda As Range, ByVal valorObjetivo As Double)

    Dim celdaValor As Double
    Dim desviacionPorcentual As Double
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
    ' Verificación
    If Not IsNumeric(celda.Value) Then Exit Sub

    If valorObjetivo = 0 Then
        celda.Interior.color = RGB(255, 0, 0)
        celda.Font.color = RGB(255, 255, 255)
        Exit Sub
    End If

    celdaValor = CDbl(celda.Value)

    ' CÁLCULO CORRECTO DE DESVIACIÓN PORCENTUAL
    ' Fórmula: |Valor - Objetivo| / Objetivo * 100
    desviacionPorcentual = Abs((celdaValor - valorObjetivo) / valorObjetivo) * 100

    ' ASIGNACIÓN DE COLOR (SIMÉTRICO para arriba y abajo)
    Select Case desviacionPorcentual
        Case 0 To 5             ' =5% desviación
            celda.Interior.color = RGB(34, 197, 94)    ' Verde (Excelente)
            celda.Font.color = RGB(0, 0, 0)
        Case 5 To 10            ' 5-10% desviación
            celda.Interior.color = RGB(251, 191, 36)   ' Amarillo (Aceptable)
            celda.Font.color = RGB(0, 0, 0)
        Case 10 To 15           ' 10-15% desviación
            celda.Interior.color = RGB(249, 115, 22)   ' Naranja (Límite)
            celda.Font.color = RGB(255, 255, 255)
        Case Is > 15            ' >15% desviación
            celda.Interior.color = RGB(220, 38, 38)    ' Rojo (Crítico)
            celda.Font.color = RGB(255, 255, 255)
    End Select
Cleanup:
    ' REACTIVAR siempre (incluso si hay error)
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
        .EnableEvents = True
        .DisplayStatusBar = True
    End With
End Sub

''================================================================================
'' NOTAS DE USO Y PERSONALIZACIÓN
''================================================================================
''
'' CÓMO PERSONALIZAR:
'' ------------------
'' 1. Colores corporativos: Modificar las constantes de COLOR al inicio del módulo
'' 2. Posición inicial: Cambiar COL_INICIO_RESULTADOS (actualmente columna H = 8)
'' 3. Ancho de cada sección: Ajustar COL_ANCHO_SECCION (por defecto = 5 columnas)
'' 4. ** SEPARACIÓN ENTRE SECCIONES **: Ajustar COL_SEPARACION_SECCIONES (por defecto = 2 columnas)
''    - Valor 1 = 1 columna vacía entre secciones (más compacto)
''    - Valor 2 = 2 columnas vacías entre secciones (más espaciado) ? RECOMENDADO
''    - Valor 3 = 3 columnas vacías entre secciones (muy espaciado)
'' 5. Espaciado vertical: Ajustar ALTO_SECCION_DATOS, ESPACIO_GRAFICOS, etc.
'' 6. Panel lateral: Modificar CrearPanelLateralInformativo()
''
'' EJEMPLO DE CONFIGURACIONES:
'' ---------------------------
''
'' ** CONFIGURACIÓN COMPACTA **
'' COL_ANCHO_SECCION = 4
'' COL_SEPARACION_SECCIONES = 1
'' Resultado: Secciones más juntas, caben más columnas en pantalla
''
'' ** CONFIGURACIÓN ESPACIADA (RECOMENDADA) **
'' COL_ANCHO_SECCION = 5
'' COL_SEPARACION_SECCIONES = 2
'' Resultado: Buena legibilidad, separación visual clara
''
'' ** CONFIGURACIÓN MUY ESPACIADA **
'' COL_ANCHO_SECCION = 6
'' COL_SEPARACION_SECCIONES = 3
'' Resultado: Máxima separación, ideal para presentaciones
''
'' DOS DISEÑOS DISPONIBLES:
'' ------------------------
'' 1. MostrarAnalisisVerticalEnHoja() - Layout con panel lateral (RECOMENDADO)
'' 2. GenerarInformeEstiloTabla() - Layout de tabla comparativa
''
'' CARACTERÍSTICAS PRINCIPALES:
'' ---------------------------
'' ? Columnas A-G protegidas (no se sobrescriben)
'' ? Resultados desde columna H en adelante
'' ? ** SEPARACIÓN AJUSTABLE ** entre cada sección de resultados
'' ? Líneas separadoras visuales opcionales (punteadas y sutiles)
'' ? Gráficos posicionados debajo de cada columna (sin superposición)
'' ? Panel lateral informativo con KPIs
'' ? Diseño profesional con colores corporativos suaves
'' ? Congelación de paneles en columna H
'' ? Formato responsive que se adapta al número de columnas
''
'' CÁLCULO DE POSICIONES:
'' ---------------------
'' Posición de columna N = COL_INICIO_RESULTADOS + ((N-1) × (COL_ANCHO_SECCION + COL_SEPARACION_SECCIONES))
''
'' Ejemplo con valores por defecto (Inicio=8, Ancho=5, Separación=2):
'' - Columna 1: 8 + (0 × 7) = Columna H (8)
'' - Columna 2: 8 + (1 × 7) = Columna O (15)
'' - Columna 3: 8 + (2 × 7) = Columna V (22)
'' - etc.
''
'' INTEGRACIÓN:
'' -----------
'' Este código se integra con tu módulo existente. Asegúrate de tener:
'' - Type EstadisticasColumna definido
'' - Función AnalizarColumna() implementada
'' - Función ObtenerProximoNumeroHoja() implementada
'' - Función CrearGraficosParaColumna() implementada
'' - UserForm sigmaproxvl con los controles correspondientes
''
''===============================================================================
'
'-------------------------------------------------------------------------------
' FUNCIONES AUXILIARES DE FORMATO
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' FUNCIÓN: FormatearSeccionTitulo
' PROPÓSITO: Formato para títulos de sección del reporte
'-------------------------------------------------------------------------------
Sub FormatearSeccionTitulo(rng As Range, titulo As String)
    With rng
        .Merge
        .Value = titulo
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Bold = True
        .Font.color = RGB(255, 255, 255)
        .Interior.color = RGB(68, 114, 196)  ' Azul medio
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Borders.ColorIndex = 1
    End With
    rng.Parent.Rows(rng.Row).RowHeight = 22
End Sub

'-------------------------------------------------------------------------------
' FUNCIÓN: FormatearEtiqueta
' PROPÓSITO: Formato para etiquetas de campos
'-------------------------------------------------------------------------------
Sub FormatearEtiqueta(rng As Range)
    With rng
        .Merge
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.Bold = True
        .Interior.color = RGB(217, 225, 242)  ' Azul muy claro
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = 48
        .IndentLevel = 1
    End With
    If rng.Parent.Rows(rng.Row).RowHeight < 18 Then
        rng.Parent.Rows(rng.Row).RowHeight = 18
    End If
End Sub

'-------------------------------------------------------------------------------
' FUNCIÓN: FormatearDato
' PROPÓSITO: Formato para celdas de datos
'-------------------------------------------------------------------------------
Sub FormatearDato(rng As Range)
    With rng
        .Merge
        .Font.Name = "Arial"
        .Font.Size = 10
        .Interior.color = RGB(255, 255, 255)  ' Blanco
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = 48
        .IndentLevel = 1
        .WrapText = True
    End With
    If rng.Parent.Rows(rng.Row).RowHeight < 18 Then
        rng.Parent.Rows(rng.Row).RowHeight = 18
    End If
End Sub
