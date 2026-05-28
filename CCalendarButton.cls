VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCalendarButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const MODULE_NAME As String = "CCalendarButton"
' 1. Declaramos un botón genérico que interceptará eventos
Public WithEvents btnDia As MSForms.CommandButton
Attribute btnDia.VB_VarHelpID = -1

' 2. Declaramos una variable para recordar quién es el UserForm que contiene al botón
Public frmPadre As Object

' ====================================================================
' EL EVENTO CENTRALIZADO: Esto ocurrirá sin importar cuál de los 42 botones se presione
' ====================================================================
Private Sub btnDia_Click()
    ' Medida de seguridad: Ignorar clics si el botón pertenece a un mes inactivo
    If btnDia.Enabled = False Then Exit Sub
    
    ' Le enviamos el texto del botón (el número del día) de regreso al formulario padre
    Call frmPadre.clicDia(btnDia.Caption)
End Sub
