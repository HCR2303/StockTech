VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SolicitudEquipo 
   Caption         =   "Solicitud de equipo"
   ClientHeight    =   6456
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5208
   OleObjectBlob   =   "SolicitudEquipo.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "SolicitudEquipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MODULE_NAME As String = "SolicitudEquipo Frm"
Private EquiposMatrix As Variant


Function GetMatrix()
    
    Dim EquiposList() As Variant
    Dim Filas As Long
    
    EquiposMatrix = DataBaseUtils.QueryTableConsult(TEquipos, TUbicaciones, CampoDB(TEquipos, equ_id_ubicaion), CampoDB(TUbicaciones, ubi_ubicacion))
        
    
    If IsEmpty(EquiposMatrix) Then
        MsgBox "Error de obteción de datos de EQUIPOS", vbCritical
        Exit Function
    End If
    
    
    Filas = UBound(EquiposMatrix, 2)
    
    ReDim EquiposList(0 To Filas)
    
    Dim i As Long
    For i = 0 To Filas
        EquiposList(i) = EquiposMatrix(1, i) & ""
    Next i
    
    Me.CodigoEquipo.List = EquiposList
    
End Function

Private Sub NuevoEquipo_Click()
    Unload Me
    NewEquip.Show
End Sub

Private Sub SCodigo_Click()
    With Me.CodigoEquipo
        
            If SCodigo.Value Then
                .Value = Empty
                .Enabled = False
                Solicitud.Visible = True
            Else
                .Value = Empty
                .Enabled = True
                Solicitud.Visible = False
            End If
   
    End With
End Sub

Private Sub SolicitarEquipo_Click()
    
    Dim tabla As String
    Dim campos As Variant
    Dim valores As Variant
    Dim user As String
    '===================================
    '   Ruta para equipos con Código
    '===================================
    If Not Me.Solicitud.Visible Then
        tabla = TSolicitudes
        campos = Array(CampoDB(TSolicitudes, sol_fecha), CampoDB(TSolicitudes, sol_usuario), CampoDB(TSolicitudes, sol_codigo), CampoDB(TSolicitudes, sol_tipo), CampoDB(TSolicitudes, sol_unidades))
        valores = Array(Now(), GetCurrentUser, CodigoEquipo.Value, "EQUIPO", 1)
        
        On Error GoTo ErrorDatos
        If DataBaseUtils.AddRegister(tabla, campos, valores) Then
        On Error GoTo 0
            MsgBox ("Registro de Equipo exitoso"), vbInformation
            Unload Me
            Exit Sub
        End If
    Else
    '===================================
    '   Ruta para equipos sin Código
    '===================================
        Unload Me
        NewEquip.Caption = "Equipo Sin Código"
        NewEquip.Show
    
    End If
    Exit Sub

ErrorDatos:
    MsgBox "Error de campos o valores en la tabla " & tabla, vbCritical
    Exit Sub
End Sub

Private Sub UserForm_Activate()
    Call GetMatrix
End Sub
Private Sub CodigoEquipo_Click()
    fila = CodigoEquipo.ListIndex
    
    If fila = -1 Or CodigoEquipo = Empty Then
        CriticidadRed.Visible = False
        CriticidadYellow.Visible = False
        CriticidadGreen.Visible = False
        Exit Sub
    End If
    
    EquipName.Caption = EquiposMatrix(2, fila) & ""
    MarcaEquipo.Caption = EquiposMatrix(3, fila) & ""
    ModeloEquipo.Caption = EquiposMatrix(4, fila) & ""
    Ubicacion.Caption = EquiposMatrix(15, fila) & ""
    
    Dim Criticidad As String
    Criticidad = LCase(EquiposMatrix(8, fila) & "")
    Select Case Criticidad
        Case "alto"
            CriticidadRed.Visible = True
            CriticidadYellow.Visible = False
            CriticidadGreen.Visible = False
        Case "medio"
            CriticidadRed.Visible = False
            CriticidadYellow.Visible = True
            CriticidadGreen.Visible = False
        Case Else
            CriticidadRed.Visible = False
            CriticidadYellow.Visible = False
            CriticidadGreen.Visible = True
    End Select
End Sub
