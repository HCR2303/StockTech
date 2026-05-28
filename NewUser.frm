VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewUser 
   Caption         =   "Añadir Nuevo Usuario"
   ClientHeight    =   2448
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "NewUser.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "NewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MODULE_NAME As String = "NewUser"
Private Sub AddNewUserBtn_Click()
    Dim controlador As control
    Dim nivel As String
    Unload AdminUsers
    
    For Each controlador In Me.Controls
        If TypeName(controlador) = "OptionButton" Then
            If controlador.Value = True Then
                nivel = controlador.Caption
            End If
        End If
    Next
    If nivel = Empty Then
        MsgBox "Selecciona un Nivel", vbExclamation
        Exit Sub
    End If
    
    Dim nombre As String
    Dim usuario As String
    Dim correo As String
    
    nombre = Me.UserName.Text
    usuario = Me.usuario.Text
    correo = Me.correo.Text
    
    If nombre = Empty Or usuario = Empty Or correo = Empty Then
        MsgBox "Todos los campos deben de estar llenos", vbExclamation
        Exit Sub
    End If
    If DataBaseUtils.LogDB("Añadir", "Usuario nuevo", "Se añade usuario: " & UCase(usuario), True) Then
        If Not DataBaseUtils.AddNewUser(nombre, usuario, correo, nivel, "contraseña") Then
            DataBaseUtils.RollBack (TAuditTrail)
        End If
    Else
        MsgBox "Operación Cancelada por el usuario", vbExclamation, "Usuario Nuevo"
    End If
    
    Unload Me
    
    
End Sub
Private Sub UserForm_Terminate()
    Unload Me
    AdminUsers.Show
End Sub
