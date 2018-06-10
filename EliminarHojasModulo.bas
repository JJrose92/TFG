Attribute VB_Name = "EliminarHojasModulo"
Sub EliminarHojas()
Attribute EliminarHojas.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MacroEliminar Macro
'

'

Dim i As Integer
Dim z As String

Application.DisplayAlerts = False

For i = 1 To 100
    z = CStr(i)
    On Error Resume Next
    Sheets(z).Delete
Next
Application.DisplayAlerts = True

End Sub
