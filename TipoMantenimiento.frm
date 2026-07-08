VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TipoMantenimiento 
   Caption         =   "Tipo de requisición"
   ClientHeight    =   2076
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2172
   OleObjectBlob   =   "TipoMantenimiento.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "TipoMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mantenimiento As String
Private Sub Requisicon_Click()
    Dim con As control
    mantenimiento = ""
    For Each con In Me.Controls
        If TypeName(con) = "OptionButton" And con.Value = True Then
            mantenimiento = con.Caption & ""
            Exit For
        End If
    Next
    If mantenimiento = "" Then
        MsgBox "Debe seleccionar el tipo de mantenimiento", vbExclamation
        Exit Sub
    End If
    Me.Hide
End Sub
