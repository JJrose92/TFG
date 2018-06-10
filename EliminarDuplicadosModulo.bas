Attribute VB_Name = "EliminarDuplicadosModulo"
Sub EliminarDuplicados(ByVal nombreHoja As String, ByVal nombreGenero As String)
Attribute EliminarDuplicados.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MacroEliminarDuplicados Macro
'

'

    'Dim nombre As String
    'nombre = ListObjects(1)
    'Worksheets(nombreHoja).Range(nombre).RemoveDuplicates
    
        Worksheets(nombreHoja).Range(nombreGenero).RemoveDuplicates Columns:=1, Header:= _
        xlYes

End Sub
