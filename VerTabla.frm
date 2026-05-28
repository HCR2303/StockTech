VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VerTabla 
   Caption         =   "Ver Tabla"
   ClientHeight    =   2772
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   1980
   OleObjectBlob   =   "VerTabla.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "VerTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MODULE_NAME As String = "VerTabla"
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
                    Case "SOLPEDs"
                        DataBaseUtils.GetExcelTable TSOLPEDs, refresh:=True
                    Case "Solicitudes"
                        DataBaseUtils.GetExcelTable TSolicitudes, refresh:=True
                    Case "Seguimiento"
                        DataBaseUtils.GetExcelTable TSOLPEDsTrack, refresh:=True
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
