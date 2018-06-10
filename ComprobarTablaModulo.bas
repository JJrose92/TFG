Attribute VB_Name = "ComprobarTablaModulo"
Function comprobarTabla(nombreTabla As String) As Boolean
    NroTablas = ActiveSheet.ListObjects.Count
    For x = 1 To NroTablas
        If ActiveSheet.ListObjects(x).Name = nombreTabla Then
            comprobarTabla = True
            Exit Function
        End If
    Next x
    crearTabla22 (nombreTabla)
    comprobarTabla = True

End Function
