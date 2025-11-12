Public Sub BorrarHojasMenosLaPrimera_Activo()
    Dim i As Integer

    Application.DisplayAlerts = False ' Evita mensajes de confirmación
    ' Recorre las hojas del libro activo desde la última hacia la segunda
    For i = ActiveWorkbook.Worksheets.count To 2 Step -1
        ActiveWorkbook.Worksheets(i).Delete
    Next i
    Application.DisplayAlerts = True ' Vuelve a activar los mensajes
End Sub
