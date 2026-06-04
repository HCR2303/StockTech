VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserLogIn 
   Caption         =   "UserForm1"
   ClientHeight    =   2244
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3036
   OleObjectBlob   =   "UserLogIn.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MODULE_NAME As String = "UserLogIn"
Public USERLOGGED As String
Public LEVELLOGGED As String
Public LOGGED As Boolean

Private Sub LogIn_Click()
    Dim pass As String
    Dim uso As String
    
    pass = Me.PasswordUser.Text
    SetCurrentUser (UCase(Me.UserName.Text))
    USERLOGGED = GetCurrentUser
    uso = Me.Caption
    
    ' Se limpia contraseña por seguridad
    Me.PasswordUser.Text = Empty
    LOGGED = Seguridad.ValidUser(USERLOGGED, pass)
    
    If LOGGED Then
        GlobalLOG = True
        If Me.Caption = "Inicio de Sesión" Then
            GlobalWHO = Seguridad.GetCurrentUser()
            If UCase(GlobalWHO) <> UCase(Environ("USERNAME")) Then
                MsgBox "Solo el propietario de la sesión de Windows puede utilizar el Sistema" & vbNewLine & _
                "Ingrese desde su propia sesión de Windows o su Computadora", vbCritical
                DataBaseUtils.resetGlobals
                Exit Sub
            End If
        Else
            GlobalWHO = UCase(Me.UserName.Text)
        End If
        Call App.StartReadEvents
        ' Si el inicio de sesión fue exitoso se cierra el formulario
        Unload Me
        
    Else
        ClearCurrentUser
        Call App.EndReadEvents
        Call DataBaseUtils.resetGlobals
    End If
    
End Sub

Private Sub Olvido_Click()
    Dim targetUser As String
    Dim targetNivel As String
    Dim respuesta As VbMsgBoxResult
    Dim privilegioAdmin As Boolean
    
    respuesta = MsgBox("Deberá solicitar ayuda de un Administrador." & vbNewLine & _
                       "¿Desea solicitar autorización ahora?", vbOKCancel + vbQuestion, "Restaurar Contraseña")
                       
    If respuesta = vbOK Then
        ' Solicitamos a quién le vamos a resetear la contraseña
        Unload Me
        LOGGED = False
        With UserLogIn
            .Caption = "Autorizar reinicio"
            .Show
        End With
        
        privilegio = Seguridad.VerificarPrivilegio(GlobalWHO, "Restaurar usuario")
        
        
        If privilegio Then
a:
            targetUser = Application.InputBox("Ingrese el usuario a restaurar:", "Restaurar Contraseña", Type:=2)
        
            ' Si se cancela el input o no se coloca nada se cancela la operación
            If targetUser = "Falso" Or targetUser = "False" Or Trim(targetUser) = "" Then
                MsgBox "Operación Cancelada"
                Exit Sub
            End If
            Dim camposUser As Variant
            camposUser = CamposTablaDB(TUsuarios)
            
            If LCase(DataBaseUtils.DatoDataBase(TUsuarios, camposUser(us_usuario), targetUser, camposUser(us_usuario))) = LCase(targetUser) Then
                If DataBaseUtils.LogDB("Reinicio", "Contraseña", "Se reestablece contraseña del usuario: " & UCase(targetUser)) Then
                    Call DataBaseUtils.UserChanges(targetUser, us_contrasenia, "contraseña")
                    MsgBox "Contraseña reestablecida correctamente", vbInformation
                    LOGGED = True
                    Exit Sub
                Else
                    App.SystemError ("Error al reestablecer contraseña")
                End If
            Else
                MsgBox "El usuario a reestablecer no existe. Verifique su sintaxis", vbExclamation, "Usuario no Encontrado"
                GoTo a
            End If
        Else
            MsgBox "No cuenta con los privilegios para esta acción", vbExclamation, "Autorización Denegada"
            
            Unload Me
        End If
    End If
End Sub

Private Sub UserForm_Click()

End Sub
