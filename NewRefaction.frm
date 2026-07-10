VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewRefaction 
   Caption         =   "Nueva Refacción"
   ClientHeight    =   3672
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "NewRefaction.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "NewRefaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private crit As Variant
Private ubic As Variant
Const MODULE_NAME As String = "NewRefaction"
Public RegistroRefaccion As Boolean
Function GetLists()
    Dim CritList() As Variant
    Dim UbicList() As Variant
    Dim Filas As Long
    
    'Comandos para adquisición de tablas de datos
    crit = DataBaseUtils.UniqueValues(ref_criticidad, TRefacciones)
    ubic = DataBaseUtils.UniqueValues(ref_ubicacion_almacen, TRefacciones)
    
    If IsEmpty(crit) Then
        'MsgBox "Error de obteción de datos de CRITICIDAD", vbCritical
        Me.Criticidad.Visible = False
        Me.CritLabel.Visible = False
    End If
    If IsEmpty(ubic) Then
        MsgBox "Error de obteción de datos de UBICACIÓN", vbCritical
    End If
    
    If Me.Criticidad.Visible Then Me.Criticidad.List = crit
    
    Ubicacion.List = ubic
    
    
End Function

Private Sub Unidades_keypress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' =======================================================
    ' INTERCEPTOR DE TECLADO PARA CAMPOS NUMÉRICOS (DECIMALES)
    ' =======================================================
    
    Select Case KeyAscii
        ' 1. Teclas Permitidas: Números del 0 al 9 (ASCII 48-57) y Retroceso (8)
        Case 48 To 57, 8
            ' Dejamos pasar la tecla con normalidad
            
        Case Else
            ' Asignar 0 a KeyAscii destruye la pulsación. El usuario no verá nada en pantalla.
            KeyAscii = 0
    End Select
End Sub
Private Sub NuevaRefaccion_Click()
    Dim ctr As control
    Dim vacio As Integer
    vacio = 0
    For Each ctr In Me.Controls
        If TypeName(ctr) = "TextBox" Or TypeName(ctr) = "ComboBox" Then
            If ctr.Value = emty And ctr.Visible = True Then vacio = vacio + 1
        End If
    Next
    If vacio > 0 Then
        MsgBox "Todos los campos son obligatorios", vbExclamation
        
        Exit Sub
    End If
    
    Dim arrRef As Variant
    Dim campos As Variant
    Dim valores As Variant
    
    If Me.Caption = "Refacción Sin Código" Then
        
        
        Dim tabla As String
        tabla = TSolicitudes
        campos = Array(CampoDB(tabla, sol_fecha), CampoDB(tabla, sol_usuario), CampoDB(tabla, sol_codigo), CampoDB(tabla, sol_tipo), CampoDB(tabla, sol_unidades), CampoDB(tabla, sol_comentarios))
        
        valores = Array(Now(), GetCurrentUser, "Solicitar Código", "REFACCIÓN", CInt(Me.Unidades.Text), "El usuario solicitó " & Me.Unidades.Text & " unidades de la REFACCIÓN: " & Me.Refaccion.Text & Chr(10) & "PRESENTACIÓN: " & Me.Presentación.Text & Chr(10) & _
                                                                            "Detalles: " & Me.Detalles.Text)
        If DataBaseUtils.LogDB("Solicitud", "Refacción Nueva: " & UCase(Me.Codigo.Text), "Se solicita Refacción nueva, Código: " & Me.Codigo.Text) Then
            If Not DataBaseUtils.AddRegister(TSolicitudes, campos, valores) Then
                MsgBox "Error al crear refacción nueva. Contacte al administrador", vbCritical, "Error de conexión"
                Unload Me
            Else
                RegistroRefaccion = True
                Unload Me
                MsgBox "Registro exitoso", vbInformation, "Registro de nueva Refacción"
                Call DataBaseUtils.GetExcelTable(TSolicitudes, refresh:=True)
                Call StringUtils.FiltrarTabla(TSolicitudes, CampoDB(tabla, sol_realizada), False)
                Call Seguridad.LockSheet(ActiveSheet)
            
            End If
        End If
        Exit Sub
    End If
    
    
    campos = Array(CampoDB(TRefacciones, ref_material), CampoDB(TRefacciones, ref_descripcion), CampoDB(TRefacciones, ref_presentacion), CampoDB(TRefacciones, ref_ubicacion_almacen), CampoDB(TRefacciones, ref_criticidad))
    valores = Array(Me.Codigo.Text, Me.Refaccion.Text, Me.Presentación.Text, Me.Ubicacion.Text, Me.Criticidad.Text)
    
    If DataBaseUtils.LogDB("Creación", "Refacción Nueva: " & UCase(Codigo.Text), "Se crea Refacción nueva, Codigo: " & Codigo.Text, True) Then
        If DataBaseUtils.AddRegister(TRefacciones, campos, valores) Then
            RegistroRefaccion = True
            If ActiveSheet.name <> TSolicitudes Then
                MsgBox "Refacción creada con éxito.", vbInformation, "Administración de Base de Datos"
                verificar = MsgBox("¿Desea verificar la información?", vbYesNo)
                
                If verificar = vbYes Then
                    Call DataBaseUtils.GetExcelTable(TRefacciones, refresh:=True)
                    Call StringUtils.SeleccionColumnasTabla(TRefacciones, campos)
                    Call Seguridad.LockSheet(ActiveSheet)
                End If
            End If
        Else
            DataBaseUtils.RollBack (TAuditTrail)
            MsgBox "Error al crear equipo nuevo. Contacte al administrador", vbCritical, "Error de conexión"
        End If
    End If
    Unload Me
    
End Sub

Private Sub UserForm_Activate()
    
    GetLists
    If Me.Caption = "Refacción Sin Código" Then
        Me.Unidades.Visible = True
        Me.UnidadesLabel.Visible = True
        Me.Codigo.Visible = False
        Me.CodigoLabel.Visible = False
        Me.Detalles.Visible = True
        Me.Detalles.Enabled = True
        Me.CritLabel.Visible = False
        Me.UbicLabel.Caption = "Más detalles"
        Me.Ubicacion.Visible = False
        Me.Ubicacion.Enabled = False
        Me.Criticidad.Visible = False
        Me.Criticidad.Enabled = False
    End If
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Me.Caption = "Refacción Sin Código" And RegistroRefaccion = False Then
        Unload Me
        SolicitudRefaccion.Show
    End If
End Sub
