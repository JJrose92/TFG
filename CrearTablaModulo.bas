Attribute VB_Name = "CrearTablaModulo"
Sub crearTabla22(nombreTabla As String)
   Application.CutCopyMode = False
   ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1"), , xlNo).Name = nombreTabla
   Range("A1").Select
   ActiveCell.FormulaR1C1 = nombreTabla
End Sub
