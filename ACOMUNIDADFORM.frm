VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ACOMUNIDADFORM 
   Caption         =   "Añadir Comunidad"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8760.001
   OleObjectBlob   =   "ACOMUNIDADFORM.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ACOMUNIDADFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub comunidadText_Change()

End Sub

Private Sub menuButton_Click()
Unload Me
MENUFORM.Show
End Sub

Private Sub okButton_Click()
    
  If crearHoja(comunidadText.Value) Then
    MsgBox "comunidadText.Value creada"
   Else
    MsgBox "La hoja estaba creada"
    
  
    
    comunidadText.Value = ""
    
End Sub

Private Sub volverButton_Click()
Unload Me
ANADIRFORM.Show
End Sub
