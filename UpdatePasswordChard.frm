VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UpdatePasswordChard 
   Caption         =   "Actualizar Contraseña"
   ClientHeight    =   3300
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2640
   OleObjectBlob   =   "UpdatePasswordChard.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UpdatePasswordChard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MODULE_NAME As String = "UpdatePasswordChard"
Public NEWPASSWORD As String
Public RESULT As Boolean
Private Sub ConfirmPass_Change()
    If Me.NewPass.Text <> Me.ConfirmPass.Text Then
        Me.ConfirmPass.ControlTipText = "Las contraseñas no coinciden"
        Me.ConfirmPass.ForeColor = vbRed
    Else
        Me.ConfirmPass.ForeColor = RGB(104, 233, 12)
    End If
End Sub

Private Sub ResetPass_Click()
    
    If Me.NewPass.Text <> Me.ConfirmPass.Text Then
        MsgBox "Las contraseñas no coinciden", vbExclamation
        Me.NewPass.Text = Empty
        Me.ConfirmPass.Text = Empty
        Exit Sub
    End If
    Dim user As String
    Dim password As String
    NEWPASSWORD = Me.NewPass.Text
    
    user = Me.UserLabel.Caption
    password = Me.CurrentPass.Text
    
    Dim validez As Boolean
    If UpdatePasswordChard.Caption = "Reboot Inicio" Then
        GoTo reinicio
    End If
    validez = Seguridad.ValidUser(user, password)
    
    If validez = False Then
        MsgBox "La contraseña actual es incorrecta", vbExclamation
        Exit Sub
    Else
        If DataBaseUtils.LogDB("Actualización", "Contraseña", "Se actualiza contraseña de usuario: " & user) Then
a:          Call DataBaseUtils.UserChanges(user, us_contrasenia, NEWPASSWORD)
            Unload Me
        Else
            MsgBox "Intente nuevamente"
            Me.NewPass.Text = Empty
            Me.ConfirmPass.Text = Empty
        End If
    End If
    Exit Sub
reinicio:
    LOGGED = True
    GlobalWHO = user
    If DataBaseUtils.SystemLogDB("Reestablecimiento", "Contraseña") Then
        GoTo a
    End If
End Sub

Private Sub UserForm_Activate()
    Dim user As String
    Dim pass As String
    user = Me.UserLabel.Caption
    pass = DataBaseUtils.DatoDataBase("Usuarios", "Usuario", user, "Contraseña")
    If pass = "AFB0CDBAC1B3C0BA3EA2" Then
        With UpdatePasswordChard
            .Caption = "Reboot Inicio"
            .UserLabel.Caption = user
            .CurrentPass.Enabled = False
            .CurrentPass.BackColor = &H80000004
            .CurrentPass.Text = Seguridad.Desencriptar(pass)
        End With
    End If
End Sub
