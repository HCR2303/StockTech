VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AuditTrail 
   Caption         =   "AuditTrail"
   ClientHeight    =   2688
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4248
   OleObjectBlob   =   "AuditTrail.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "AuditTrail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MODULE_NAME As String = "AuditTrail"
Public Resultado As String

Private Sub CancelAudit_Click()
    GlobalLOG = False
    Unload Me
End Sub

Private Sub UserForm_Activate()
    Me.Comentario.Text = Empty
End Sub

' Interceptamos si el usuario intenta cerrar con la "X" de la ventana
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then ' 0 = El usuario presionó la "X"
        GlobalLOG = False
    End If
End Sub

Private Sub LogAudit_Click()
    If Trim(Me.Comentario.Value) = "" Then
        MsgBox "Debe ingresar una justificación.", vbExclamation
        Exit Sub
    End If
    
    Resultado = Me.Comentario.Value
    Me.Hide
End Sub
