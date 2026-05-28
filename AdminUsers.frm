VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AdminUsers 
   Caption         =   "Administración de Usuarios"
   ClientHeight    =   3396
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "AdminUsers.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "AdminUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private UserMatrix As Variant
Const MODULE_NAME As String = "AdminUsers Frm"

Private Sub AddNewUser_Click()
    Unload Me
    NewUser.Show
    
End Sub

Private Sub ContraseñaUser_Click()
    If ListUsers.ListIndex = -1 Then
        MsgBox "Selecciona el usuario", vbExclamation
        Exit Sub
    End If
    UpdatePasswordChard.UserLabel.Caption = ListUsers.Value
    UpdatePasswordChard.Show
    If UpdatePasswordChard.RESULT Then
        Unload Me
    End If
End Sub

Private Sub DeleteUser_Click()
    If ListUsers.ListIndex = -1 Then
        MsgBox "Selecciona el usuario", vbExclamation
        Exit Sub
    End If
    Dim usuario As String
    
    usuario = ListUsers.Value
    Unload AdminUsers
    nivel = LevelUser.Caption
    
    If nivel = "Administrador" Then
        If Seguridad.EsDesarrollador = False Then
            MsgBox "La eliminación de este usuario solo puede ser realizada por el administrador operativo", vbCritical, "No Autorizado"
            Exit Sub
        End If
    End If
    
    If DataBaseUtils.LogDB("Eliminación", "Usuario: " & usuario, "Se elimina usuario " & usuario, True) Then
        If Not DataBaseUtils.DeleteUser(usuario) Then
            DataBaseUtils.RollBack (TAuditTrail)
        End If
    Else
        MsgBox "Operación cancelada", vbExclamation, "Eliminación de usuario"
    End If
    AdminUsers.Show
    Exit Sub
           
End Sub

Private Sub ListUsers_Click()
    fila = ListUsers.ListIndex
    If fila = -1 Then Exit Sub

    UserName.Caption = UserMatrix(1, fila) & ""
    LevelUser.Caption = UserMatrix(4, fila) & ""
    
End Sub


Private Sub UserForm_Activate()
    RefreshList
End Sub
Function RefreshList()
    Dim listaNombres() As Variant
    Dim i As Long
    Dim totalFilas As Long
    Dim colDeseada As Integer
    
    UserMatrix = DataBaseUtils.TableDataBase("Usuarios", "*")
    
    If IsEmpty(UserMatrix) Then
        MsgBox "Error de conexión", vbCritical
        Exit Function
    End If
    
    colDeseada = 2
    If UBound(UserMatrix, 1) < colDeseada Then
        MsgBox "La columna " & colDeseada & " no existe en la matriz.", vbCritical
        Exit Function
    End If
    
    totalFilas = UBound(UserMatrix, 2)
    
    ReDim listaNombres(0 To totalFilas)
    
    For i = 0 To totalFilas
        listaNombres(i) = UserMatrix(colDeseada, i) & ""
    Next
    
    ListUsers.List = listaNombres
End Function
