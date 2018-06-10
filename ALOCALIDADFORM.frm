VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ALOCALIDADFORM 
   Caption         =   "Añadir Localidad"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8760.001
   OleObjectBlob   =   "ALOCALIDADFORM.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ALOCALIDADFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComunidadList_Click()

End Sub

Private Sub menuButton_Click()
Unload Me
MENUFORM.Show
End Sub

Private Sub okButton_Click()
Dim comunidad As String

comunidad = ComunidadList.Value
Worksheets(comunidad).Activate
Range("B1:B5").Value = nombreText.Value
nombreText.Value = ""

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub volverButton_Click()
Unload Me
ANADIRFORM.Show
End Sub


Private Sub UserForm_Initialize()
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ComunidadList.AddItem (ws.Name)
    Next ws
End Sub

