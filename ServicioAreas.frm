VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ServicioAreas 
   Caption         =   "Servicio para Áreas"
   ClientHeight    =   3804
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5208
   OleObjectBlob   =   "ServicioAreas.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ServicioAreas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private UbicacionesMatrix As Variant
Private EquiposMatrix As Variant
Private ProveedorMatrix As Variant
Private RefaccionMatrix As Variant
Private IdSolicitud As Integer
Private TipoSolicitud As String
Const MODULE_NAME As String = "ServicioAreas FORM"
Function GetMatrix()
    Dim UbicacionesList() As Variant
    
    Dim Filas As Long
    
    'Comandos para adquisición de tablas de datos
    
    Dim camposUbic As Variant
    
    camposUbic = CamposTablaDB(TUbicaciones)
    
    UbicacionesMatrix = DataBaseUtils.TableDataBase(TUbicaciones, "*")
    
    If IsEmpty(UbicacionesMatrix) Then
        MsgBox "Error de obteción de datos de UBICACIONES", vbCritical
        
    End If
    
    'Comandos para adquisición de lista de UBICACIONES
    Filas = UBound(UbicacionesMatrix, 2)
    ReDim UbicacionesList(0 To Filas)
    For i = 0 To Filas
        UbicacionesList(i) = UbicacionesMatrix(4, i) & ""
    Next
    UbicacionCodigo.List = UbicacionesList
    
End Function


Private Sub SolicitarRefaccion_Click()
    If CodigoUbic.Caption = "" Then
        MsgBox "Debe elegir una ubicación válida", vbExclamation
        Exit Sub
    End If
    Dim campos As Variant
    Dim valores As Variant
    
    campos = CamposTablaDB(TSolicitudes)
    StringUtils.DropFromArray campos, 0
    StringUtils.DropFromArray campos, UBound(campos)
    Dim camposVisibles As Variant
    camposVisibles = campos
    campos = Array(CampoDB(TSolicitudes, sol_fecha), CampoDB(TSolicitudes, sol_usuario), CampoDB(TSolicitudes, sol_codigo), CampoDB(TSolicitudes, sol_tipo), CampoDB(TSolicitudes, sol_mantenimiento))
    valores = Array(Now, UCase(GetCurrentUser), Me.CodigoUbic.Caption, "Servicio Área", StringUtils.SetMantenimiento)
    If DataBaseUtils.LogDB("Registro de solicitud", "Servicio de Área", "Se solicita servició para el área " & UCase(NombreUbic.Caption)) Then
        If DataBaseUtils.AddRegister(TSolicitudes, campos, valores) Then
            MsgBox "Registro de solicitud exitosa"
            Call DataBaseUtils.GetExcelTable(TSolicitudes, refresh:=True)
            Call StringUtils.SeleccionColumnasTabla(TSolicitudes, camposVisibles)
            Call Seguridad.LockSheet(ActiveSheet)
            Unload Me
        Else
            MsgBox "Ha ocurrido un error de registro", vbExclamation
            DataBaseUtils.RollBack (TAuditTrail)
        End If
    End If
    
End Sub

Private Sub UbicacionCodigo_Change()
    Dim fila As Long
    fila = UbicacionCodigo.ListIndex
    
    ' =======================================================
    ' COINCIDENCIA EXACTA
    ' =======================================================
    
    If fila <> -1 And Trim(UbicacionCodigo.Value) <> "" Then
        
        If UCase(UbicacionCodigo.Value) = UCase(UbicacionCodigo.List(fila)) Then
            ' Solo si hay coincidencia perfecta, disparamos el motor
            Call RefreshUbicacion
            Exit Sub
        End If
        
    End If
    
    ' =======================================================
    ' ESTADO DE ESPERA (Limpieza Visual)
    ' =======================================================
    ' Si el usuario sigue escribiendo o no hay coincidencia exacta,
    ' enviamos una bandera True para limpiar la pantalla y evitar fantasmas.
    Call RefreshUbicacion(LimpiarPantalla:=True)
End Sub
Private Sub UserForm_Activate()
    GetMatrix
End Sub
Function RefreshUbicacion(Optional LimpiarPantalla As Boolean = False)
    Dim fila As Long
        
    ' =======================================================
    ' MODO LIMPIEZA (Señal de Espera)
    ' =======================================================
    If LimpiarPantalla Then
        CodigoUbic.Caption = ""
        NombreUbic.Caption = ""
        Exit Function
    End If
    
    ' =======================================================
    ' MODO EXTRACCIÓN Y RENDERIZADO
    ' =======================================================
    fila = UbicacionCodigo.ListIndex
    
    ' Inyección segura a los Labels (previene Error 13 por nulos)
    NombreUbic.Caption = UbicacionesMatrix(ubi_ubicacion, fila) & ""
    CodigoUbic.Caption = UbicacionesMatrix(ubi_id_ubicacion, fila) & " "
    
        
End Function
