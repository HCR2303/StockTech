VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TipoServicio 
   Caption         =   "Tipo de Servicio"
   ClientHeight    =   1836
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2172
   OleObjectBlob   =   "TipoServicio.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "TipoServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Area_Click()
    Unload Me
    ServicioAreas.Show
End Sub

Private Sub Equipo_Click()
    Unload Me
    SolicitudEquipo.Caption = "Servicio a Equipo"
    SolicitudEquipo.Show
End Sub
