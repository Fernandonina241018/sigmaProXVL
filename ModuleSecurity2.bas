Sub ProtegerTodasLasHojasYEstructura()
    Dim ws As Worksheet
    
    ' Bloquear la estructura del libro para evitar que se agreguen/eliminen hojas
    ActiveWorkbook.Protect Structure:=True, Windows:=False, password:=VAL123
    
    ' Recorrer todas las hojas del libro
    For Each ws In ActiveWorkbook.Worksheets
        ' Desproteger la hoja por si ya tenía protección
        ws.Unprotect password:=VAL123
        
        ' Bloquear la hoja con las opciones de protección de celdas y objetos
        ws.Protect password:=VAL123, Contents:=True, _
                     Scenarios:=True, UserInterfaceOnly:=False
    Next ws
    
    ' Si usaste una subrutina para desproteger la estructura, puedes llamarla aquí
    ' para asegurarte de que está protegida de nuevo.
End Sub

Sub ProtegerTodasLasHojasYEstructura2()
    Dim ws As Worksheet
    
    ' Bloquear la estructura del libro para evitar que se agreguen/eliminen hojas
    ActiveWorkbook.Protect Structure:=False, Windows:=False, password:=VAL123
    
    ' Recorrer todas las hojas del libro
    For Each ws In ActiveWorkbook.Worksheets
        ' Desproteger la hoja por si ya tenía protección
        ws.Unprotect password:=VAL123
        
        ' Bloquear la hoja con las opciones de protección de celdas y objetos
        'ws.Protect password:=123, Contents:=True, _
                     Scenarios:=True, UserInterfaceOnly:=False
    Next ws
    
    ' Si usaste una subrutina para desproteger la estructura, puedes llamarla aquí
    ' para asegurarte de que está protegida de nuevo.
End Sub

