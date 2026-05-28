VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TipoSolicitud 
   Caption         =   "Crear solicitud"
   ClientHeight    =   1788
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2196
   OleObjectBlob   =   "TipoSolicitud.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "TipoSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MODULE_NAME As String = "TipoSolicitud"

Private Sub SolEquipo_Click()
    
    If Me.Caption = "Códigos Nvos" Then
        Unload Me
        NewEquip.Show
        Exit Sub
    End If
    Unload Me
    SolicitudEquipo.Show
End Sub

Private Sub SolRefaccion_Click()
    
    If Me.Caption = "Códigos Nvos" Then
        Unload Me
        NewRefaction.Show
        Exit Sub
    End If
    Unload Me
    SolicitudRefaccion.Show
End Sub
