Attribute VB_Name = "ComprobarHojaModulo"
Sub comprobarHoja(nombreHoja As String)
For i = 1 To Worksheets.Count
        If Worksheets(i).Name = nombreHoja Then
            Exit Sub
        End If
    Next
    crearHoja22 (nombreHoja)
End Sub
