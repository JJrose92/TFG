Attribute VB_Name = "InsertarGeneroModulo"
Sub insertarGenero(ByVal nombreHoja As String, ByVal nombreGenero As String)
    'Dim nombre As String
    'Dim nombreGenero As String
    'nombreGenero = ""
    'nombre = "Genero"
    comprobarHoja (nombreHoja)
    Sheets(nombreHoja).Activate
    comprobarTabla (nombreHoja)
    insertarNombreTabla nombreGenero, nombreHoja
    Columns("A:A").Select
    Selection.Columns.AutoFit
    OrdenarTabla (nombreHoja)
    
    'EliminarDuplicados nombreHoja, nombreGenero
    
        
End Sub
