VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewEquip 
   Caption         =   "Equipo Nuevo"
   ClientHeight    =   3804
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "NewEquip.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "NewEquip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private crit As Variant
Private ubic As Variant
Const MODULE_NAME As String = "NewEquip"
Public RegistroEquipo As Boolean

Function GetLists()
    Dim CritList() As Variant
    Dim UbicList() As Variant
    Dim Filas As Long
    
    'Comandos para adquisición de tablas de datos
    crit = DataBaseUtils.UniqueValues(equ_criticidad, TEquipos)
    ubic = DataBaseUtils.UniqueValues(equ_id_ubicacion, TEquipos)
    
    If IsEmpty(crit) Then
        MsgBox "Error de obteción de datos de CRITICIDAD", vbCritical
        Exit Function
    End If
    If IsEmpty(ubic) Then
        MsgBox "Error de obteción de datos de UBICACIÓN", vbCritical
        Exit Function
    End If
    
    Criticidad.List = crit
    
    Ubicacion.List = ubic
    
    
End Function
Private Sub NuevoEquipo_Click()
    Dim camposEqui As Variant
    Dim campos As Variant
    Dim valores As Variant
    
    
    
    Dim requisito As control
    cont = 0
    For Each requisito In Controls
        If TypeName(requisito) = "ComboBox" Or TypeName(requisito) = "TextBox" Then
            texto = Trim(requisito.Text)
            If texto = Empty And requisito.Enabled = True Then
                cont = cont + 1
            End If
        End If
    Next
    
    If cont > 0 Then
    
        MsgBox "Todos los campos deben ser llenados, intente de nuevo", vbInformation
        Exit Sub
    End If
    
    If Me.Caption = "Equipo Sin Código" Then
        ' Al realizar un registro de un equipo sin código, Los comentarios se vuelven las especificaciones a solicitar
        Dim tabla As String
        tabla = TSolicitudes
        campos = Array(CampoDB(tabla, sol_fecha), CampoDB(tabla, sol_usuario), CampoDB(tabla, sol_codigo), CampoDB(tabla, sol_tipo), CampoDB(tabla, sol_comentarios))
        
        valores = Array(Now(), GetCurrentUser, "Solicitar Código", "EQUIPO", "El usuario solicitó un EQUIPO: " & Me.Equipo.Text & Chr(10) & "MARCA: " & Me.Marca.Text & Chr(10) & _
                                                                            "MODELO: " & Me.Modelo.Text & Chr(10) & "No. SERIE: " & Me.Serie.Text & Chr(10) & "UBICACIÓN: " & _
                                                                            "Detalles: " & Me.Detalles.Text)
        If DataBaseUtils.LogDB("Solicitud", "Equipo Nuevo: " & Codigo.Text, "Se solicita equipo nuevo, CÓDIGO: " & Codigo.Text, True) Then
            If Not DataBaseUtils.AddRegister(TSolicitudes, campos, valores) Then
                DataBaseUtils.RollBack (TAuditTrail)
                MsgBox "Error al crear equipo nuevo. Contacte al administrador", vbCritical, "Error de conexión"
                Unload Me
            Else
                RegistroEquipo = True
                Unload Me
                MsgBox "Registro exitoso", vbInformation, "Registro de nuevo Equipo"
            End If
        End If
        Exit Sub
    End If
    
    camposEqui = CamposTablaDB(TEquipos)
    campos = Array(camposEqui(equ_codigo), camposEqui(equ_nombre), camposEqui(equ_marca), camposEqui(equ_modelo), camposEqui(equ_no_serie), _
                camposEqui(equ_id_ubicacion), camposEqui(equ_criticidad))
    valores = Array(Me.Codigo.Text, Me.Equipo.Text, Me.Marca.Text, Me.Modelo.Text, Me.Serie.Text, Me.Ubicacion.Text, Me.Criticidad.Text)
    If DataBaseUtils.LogDB("Creación", "Equipo Nuevo: " & UCase(Codigo.Text), "Se crea equipo nuevo, Codigo: " & Codigo.Text, True) Then
        If DataBaseUtils.AddRegister(TEquipos, campos, valores) Then
            RegistroEquipo = True
            MsgBox "Equipo creado con éxito.", vbInformation, "Administración de Base de Datos"
            verificar = MsgBox("¿Desea verificar la información?", vbYesNo)
            
            If verificar = 6 Then
                MsgBox "Se mostrará la tabla de datos de Equipos a continuación... ", vbInformation
                Call DataBaseUtils.GetExcelTable(TEquipos, refresh:=True)
                Call StringUtils.SeleccionColumnasTabla(TEquipos, campos, True)
            End If
        Else
            DataBaseUtils.RollBack (TAuditTrail)
            MsgBox "Error al crear equipo nuevo. Contacte al administrador", vbCritical, "Error de conexión"
        End If
    End If
    Unload Me
    
End Sub
Private Sub UserForm_Activate()
    Call GetLists
    If Me.Caption = "Equipo Sin Código" Then
        Me.Codigo.Enabled = False
        Me.Codigo.Visible = False
        Me.CodigoLabel.Visible = False
        Me.Criticidad.Visible = False
        Me.Criticidad.Enabled = False
        Me.Ubicacion.Visible = False
        Me.Ubicacion.Enabled = False
        Me.UbicLabel.Visible = False
        Me.CritLabel.Caption = "Más Detalles"
        Me.Detalles.Visible = True
        Me.Detalles.Enabled = True
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Me.Caption = "Equipo Sin Código" And RegistroEquipo = False Then
        Unload Me
        SolicitudEquipo.Show
    End If
End Sub
