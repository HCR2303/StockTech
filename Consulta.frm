VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Consulta 
   Caption         =   "Consulta en DB"
   ClientHeight    =   2724
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   1896
   OleObjectBlob   =   "Consulta.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MODULE_NAME As String = "Consulta"
Private Sub Consultar_Click()
    Dim ctrl As control ' Corregido el tipado
    Dim tablaSeleccionada As String
    Dim valido As Boolean
    ' Recorremos los controles del formulario
    For Each ctrl In Me.Controls
        ' 1. Verificamos que sea un OptionButton
        If TypeOf ctrl Is MSForms.OptionButton Then
            
            ' 2. ¡CRÍTICO! Solo nos importa el que está seleccionado
            If ctrl.Value = True Then
                valido = True
                Select Case ctrl.Name
                    Case "Equipos"
                        DataBaseUtils.GetExcelTable TEquipos, refresh:=True
                        
                    Case "Refacciones"
                        DataBaseUtils.GetExcelTable TRefacciones, refresh:=True
                        
                    Case "Presupuestos"
                        DataBaseUtils.GetExcelTable TPresupuestos, refresh:=True
                        
                    Case "Proveedores"
                        DataBaseUtils.GetExcelTable TProveedores, refresh:=True
                End Select
                Exit For
            End If
        End If
    Next ctrl
    Call App.EnfocarStockTechLabels
    If Not valido Then
        MsgBox "Debe seleccionar al menos una opción", vbExclamation, "Sin Selección"
        Exit Sub
    End If
    Call Seguridad.LockSheet(ActiveSheet)
    Unload Me
End Sub

