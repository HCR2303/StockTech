VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserDetails 
   Caption         =   "Detalles de Usuario"
   ClientHeight    =   3420
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4224
   OleObjectBlob   =   "UserDetails.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private UserMatrix As Variant
Const MODULE_NAME As String = "UserDetails Frm"

Private Sub GetDetails_Click()
    fila = ListUsers.ListIndex
    If fila = -1 Then
        MsgBox "Debe seleccionar un usuario"
        Exit Sub
    End If
    Dim user As String
    user = CStr(ListUsers.Value)
    Call AccionesRibbon.VerUserDetails(user)
    
    Unload Me
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
