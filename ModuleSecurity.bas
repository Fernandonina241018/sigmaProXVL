'---------------------------------------------------------------------------------------
' Constante de contraseña
'---------------------------------------------------------------------------------------
Private Const Contrasena As String = "Dh@ra/24-1018/fanr"

'---------------------------------------------------------------------------------------
' Procedure : MostrarError
' Purpose   : Muestra un mensaje de error estándar
'---------------------------------------------------------------------------------------
Private Sub MostrarError(ByVal mensaje As String)
    MsgBox "? " & mensaje, vbCritical, "Error"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DesprotegerLibroYHoja
' Purpose   : Desprotege la estructura del libro y la hoja activa.
' Esta linea es una prueba para verificar el Update desde Github
'---------------------------------------------------------------------------------------
Public Sub DesprotegerLibroYHoja()
    On Error GoTo ErrorHandler

    ActiveWorkbook.Unprotect password:=Contrasena
    ActiveSheet.Unprotect password:=Contrasena

    If Not ActiveWorkbook.ProtectStructure And Not ActiveSheet.ProtectContents Then
        Debug.Print "? La hoja '" & ActiveSheet.Name & "' y la estructura del libro han sido desbloqueadas."
    Else
        MsgBox "?? No se pudo desbloquear la hoja o la estructura del libro. Verifique la contraseña.", vbExclamation
    End If
    Exit Sub

ErrorHandler:
    MostrarError "Error al intentar desproteger el libro o la hoja."
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ProtegerLibroYHoja
' Purpose   : Protege la hoja especificada y la estructura del libro.
'---------------------------------------------------------------------------------------
Public Sub ProtegerLibroYHoja(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler

    ws.Protect password:=Contrasena, Contents:=True, Scenarios:=True, UserInterfaceOnly:=False, DrawingObjects:=False
    ActiveWorkbook.Protect password:=Contrasena, Structure:=True, Windows:=False

    If ActiveWorkbook.ProtectStructure And ws.ProtectContents Then
        Debug.Print "?? La hoja '" & ws.Name & "' y la estructura del libro han sido protegidas."
    Else
        MsgBox "?? No se pudo proteger la hoja o la estructura del libro.", vbExclamation
    End If
    Exit Sub

ErrorHandler:
    MostrarError "Error al intentar proteger el libro o la hoja."
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Asegúrate de que el UserForm esté cargado y visible
    If UserData.Visible Then
        ' Desactiva los eventos para evitar un bucle infinito
        Application.EnableEvents = False
        
        ' Verifica si la selección no es de una sola celda
        If Target.Cells.count > 1 Then
            ' Rellena el ComboBox del rango de inicio con la primera celda
            UserData.ComboBox1.Value = Target.Cells(1, 1).Address(False, False)
            
            ' Rellena el ComboBox del rango final con la última celda
            UserData.ComboBox2.Value = Target.Cells(Target.Cells.count).Address(False, False)
        End If
        
        ' Reactiva los eventos
        Application.EnableEvents = True
    End If
End Sub

