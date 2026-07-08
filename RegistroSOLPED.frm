VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegistroSOLPED 
   Caption         =   "Registro de Nueva SOLPED"
   ClientHeight    =   9360.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10272
   OleObjectBlob   =   "RegistroSOLPED.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "RegistroSOLPED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private CuentasMatrix As Variant
Private EquiposMatrix As Variant
Private ProveedorMatrix As Variant
Private RefaccionMatrix As Variant
Private UbicacionesMatrix As Variant
Private IdSolicitud As Integer
Private TipoSolicitud As String
Const MODULE_NAME As String = "RegistroSOLPED FORM"

Function GetMatrix()
    Dim CuentasList() As Variant
    Dim EquiposList() As Variant
    Dim ProveedorList() As Variant
    Dim RefaccionList() As Variant
    Dim UbicacionesList() As Variant
    Dim Filas As Long
    
    
    
    
    'Comandos para adquisición de lista de UBICACIONES
    
    'Comandos para adquisición de tablas de datos
    Dim camposCuent As Variant
    Dim camposEqu As Variant
    Dim camposUbic As Variant
    camposCuent = CamposTablaDB(TCuentas)
    camposEqu = CamposTablaDB(TEquipos)
    camposUbic = CamposTablaDB(TUbicaciones)
    UbicacionesMatrix = DataBaseUtils.TableDataBase(TUbicaciones, "*")
    CuentasMatrix = DataBaseUtils.QueryTableConsult(TPresupuestos, TCuentas, camposCuent(cta_id_cuenta), camposCuent(cta_nombre))
    EquiposMatrix = DataBaseUtils.QueryTableConsult(TEquipos, TUbicaciones, camposUbic(ubi_id_ubicacion), camposUbic(ubi_ubicacion))
    ProveedorMatrix = DataBaseUtils.TableDataBase(TProveedores, "*")
    RefaccionMatrix = DataBaseUtils.TableDataBase(TRefacciones, "*")
    
    If IsEmpty(UbicacionesMatrix) Then
        MsgBox "Error de obteción de datos de UBICACIONES", vbCritical
    End If
    
    If IsEmpty(CuentasMatrix) Then
        MsgBox "Error de obteción de datos de CUENTAS", vbCritical
        
    End If
    If IsEmpty(EquiposMatrix) Then
        MsgBox "Error de obteción de datos de EQUIPOS", vbCritical
        
    End If
    If IsEmpty(ProveedorMatrix) Then
        MsgBox "Error de obteción de datos de PROVEEDORES", vbCritical
        
    End If
    If IsEmpty(RefaccionMatrix) Then
        MsgBox "Error de obteción de datos de REFACCIONES", vbCritical
        
    End If
    
    
    'Comandos para adquisición de lista de CUENTAS
    Filas = UBound(CuentasMatrix, 2)
    ReDim CuentasList(0 To Filas)
    For i = 0 To Filas
        CuentasList(i) = CuentasMatrix(2, i) & ""
    Next
    Numero.List = CuentasList
    
    
    'Comandos para adquisición de lista de EQUIPOS
    Filas = UBound(EquiposMatrix, 2)
    ReDim EquiposList(0 To Filas)
    For i = 0 To Filas
        EquiposList(i) = EquiposMatrix(2, i) & ""
    Next
    Me.EquipName.List = EquiposList
    
    'Comandos para adquisición de lista de PROVEEDORES
    Filas = UBound(ProveedorMatrix, 2)
    ReDim ProveedorList(0 To Filas)
    For i = 0 To Filas
        ProveedorList(i) = ProveedorMatrix(2, i) & ""
    Next
    Me.ProveedorName.List = ProveedorList
    
    'Comandos para adquisición de lista de REFACCIONES
    Filas = UBound(RefaccionMatrix, 2)
    ReDim RefaccionList(0 To Filas)
    For i = 0 To Filas
        RefaccionList(i) = RefaccionMatrix(2, i) & ""
    Next
    Me.RefaccionName.List = RefaccionList
    
    'Comandos para adquisición de lista de UBICACIONES
    Filas = UBound(UbicacionesMatrix, 2)
    ReDim UbicacionesList(0 To Filas)
    For i = 0 To Filas
        UbicacionesList(i) = UbicacionesMatrix(4, i) & ""
    Next
    UbicacionCodigo.List = UbicacionesList
    
End Function
Private Sub AreaOption_Change()
    AreaOption_Click
End Sub

Private Sub AreaOption_Click()
    If Me.AreaOption.Value Then
        Me.Area.Enabled = True
        Me.Area.Visible = True
        Me.Equipo.Enabled = False
        Me.Refaccion.Enabled = False
        Me.Equipo.Visible = False
        Me.Refaccion.Visible = False
    Else
        Me.Area.Enabled = False
        Me.Equipo.Enabled = False
        Me.Refaccion.Enabled = False
        Me.Equipo.Visible = False
        Me.Refaccion.Visible = False
    End If
End Sub




Private Sub EquipOption_Change()
    EquipOption_Click
End Sub

Private Sub EquipOption_Click()
    If Me.EquipOption.Value Then
        Me.Equipo.Enabled = True
        Me.Equipo.Visible = True
        Me.ServicioEquipo.Visible = True
        Me.ServicioEquipo.Enabled = True
        Me.Area.Enabled = False
        Me.Refaccion.Enabled = False
        Me.Area.Visible = False
        Me.Refaccion.Visible = False
        
    Else
        Me.Equipo.Enabled = False
        Me.Area.Enabled = False
        Me.Refaccion.Enabled = False
        Me.Area.Visible = False
        Me.Refaccion.Visible = False
    End If
End Sub

Private Sub RefaccionName_Change()
    Dim fila As Long
    fila = RefaccionName.ListIndex
    
    ' =======================================================
    ' ESCUDO DE COINCIDENCIA EXACTA
    ' =======================================================
    ' Verificamos que el índice exista y que la caja no esté vacía.
    If fila <> -1 And Trim(RefaccionName.Value) <> "" Then
        
        ' Comparamos el texto escrito con el texto real del elemento en la lista.
        If UCase(RefaccionName.Value) = UCase(RefaccionName.List(fila)) Then
            ' La coincidencia es absoluta. Procedemos a renderizar.
            Call RefreshRefaction
            Exit Sub
        End If
        
    End If
    
    ' =======================================================
    ' ESTADO DE ESPERA (Limpieza Visual)
    ' =======================================================
    ' Si no hay coincidencia exacta (el usuario sigue escribiendo),
    ' mandamos la señal de limpiar la interfaz.
    Call RefreshRefaction(LimpiarPantalla:=True)
End Sub

' Mantenemos el Click como redundancia de seguridad para el uso del ratón.
Private Sub RefaccionName_Click()
    Call RefaccionName_Change
End Sub
Function RefreshRefaction(Optional UltimoDato As Boolean = False, Optional LimpiarPantalla As Boolean = False)
    Dim fila As Long
    Dim Criticidad As String
    
    ' =======================================================
    ' MODO LIMPIEZA (Señal de Espera)
    ' =======================================================
    If LimpiarPantalla Then
        Me.RefaccionCodigo.Caption = ""
        Stock.Caption = ""
        
        Rrojo.Visible = False
        Ramarillo.Visible = False
        Rverde.Visible = False
        Nulo.Visible = False
        Exit Function
    End If
    
    ' =======================================================
    ' MODO EXTRACCIÓN Y RENDERIZADO
    ' =======================================================
    fila = RefaccionName.ListIndex
    
    If UltimoDato Then
        fila = CInt(UBound(RefaccionMatrix, 2))
    End If
    
    ' Inyección segura a los Labels (previene Error 13 por nulos)
    Me.RefaccionCodigo.Caption = RefaccionMatrix(1, fila) & ""
    Stock.Caption = RefaccionMatrix(4, fila) & " " & RefaccionMatrix(3, fila)
    If UCase(RefaccionName.Value) Like "SER" & "*" Then
        Me.Unidades.Visible = False
        Me.UnidadesLabel.Visible = False
    End If
    ' =======================================================
    ' SEMÁFORO DE CRITICIDAD Y BLINDAJE DE TEXTO
    ' =======================================================
    ' Aplicamos LCase y Trim para estandarizar el texto y evitar que
    ' un espacio accidental (ej. "alto ") rompa la lectura.
    Criticidad = LCase(Trim(RefaccionMatrix(6, fila) & ""))
    
    Select Case Criticidad
        Case "alto"
            Rrojo.Visible = True
            Ramarillo.Visible = False
            Rverde.Visible = False
            Nulo.Visible = False
            
        Case "medio"
            Rrojo.Visible = False
            Ramarillo.Visible = True
            Rverde.Visible = False
            Nulo.Visible = False
            
        Case "bajo"
            Rrojo.Visible = False
            Ramarillo.Visible = False
            Rverde.Visible = True
            Nulo.Visible = False
            
        Case Else
            ' Se activa el estado Nulo y se apagan estrictamente los demás
            Rrojo.Visible = False
            Me.Ramarillo.Visible = False
            Rverde.Visible = False
            Nulo.Visible = True
    End Select
    
End Function
Private Sub RefactionOption_Change()
    RefactionOption_Click
End Sub

Private Sub RefactionOption_Click()
    If Me.RefactionOption.Value Then
        Me.Refaccion.Enabled = True
        Me.Refaccion.Visible = True
        Me.Area.Enabled = False
        Me.Area.Visible = False
        Me.Equipo.Visible = True
        Me.Equipo.Enabled = True
        Me.ServicioEquipo.Visible = False
        Me.ServicioEquipo.Enabled = False
    Else
        Me.Refaccion.Enabled = False
        Me.Area.Enabled = False
        Me.Equipo.Enabled = False
        Me.Area.Visible = False
        Me.Equipo.Visible = False
    End If
End Sub

Private Sub RegisSOLPED_Click()
    Const PROC_NAME As String = "RegisSOLPED_Click"
    
    Dim espacio As control
    Dim nuloCount As Integer
    Dim r As VbMsgBoxResult
    
    ' Manejo de errores formal
    On Error GoTo ErrorHandler

    '============================================================
    ' 1. Validación de interfaz
    '============================================================
    nuloCount = 0
    For Each espacio In Me.Controls
        If espacio.Enabled Then
            If TypeOf espacio Is MSForms.ComboBox Or TypeOf espacio Is MSForms.TextBox Then
                If Trim(espacio.Text) = vbNullString Then
                    If espacio.Parent.Visible = True Then ' ProveedorCodigo es opcional según lógica previa
                        If espacio.name <> "ProveedorCodigo" And espacio.Visible = True Then
                            nuloCount = nuloCount + 1
                        End If
                    End If
                End If
            End If
        End If
    Next espacio

    If nuloCount > 0 Then
        MsgBox "Uno o más campos obligatorios requieren ser llenados para realizar el registro.", _
               vbInformation, "Validación StockTech"
        Exit Sub
    End If

    '============================================================
    ' 2. Lógica de Negocio y Base de Datos
    '============================================================
    
    ' Intentamos obtener el ID correlativo
    Dim posibleError As Boolean
    posibleError = getIdSolicitud
    If posibleError Then
        If IdSolicitud = Empty Or IdSolicitud = 0 Then
            Exit Sub
        End If
    End If

    ' Intentamos insertar la SOLPED
    ' Nota: AddSOLPED debe internamente llamar a AddRegister
    If AddSOLPED(IdSolicitud) Then
        
        MsgBox "Se ha realizado un registro exitoso.", vbInformation, "Éxito"
        Dim campoSegSPD As Variant
        campossegspd = CamposTablaDB(TSOLPEDsTrack)
        
        '---PUNTO COMMIT----
        If TipoSolicitud = "Automática" Then
            If DataBaseUtils.LogDB("Registro de solicitud automática", "ID de Solicitud: " & CStr(IdSolicitud), "Se crea solicitud automática para SOLPED: " & Me.SOLPED.Text) = False Then
                MsgBox "Trazabilidad comprometida. Contacte al administrador", vbCritical
                App.SystemError ("Error de registro de registro en AuditTrail")
                Exit Sub
            End If
        End If
        If DataBaseUtils.LogDB("Registro de SOLPED", "SOLPED: " & Me.SOLPED.Text, "Se crea SOLPED: " & Me.SOLPED.Text & "con Id de solicitud: " & IdSolicitud) = False Then
            MsgBox "Trazabilidad comprometida. Contacte al administrador", vbCritical
            App.SystemError ("Error de registro de registro en AuditTrail")
            Exit Sub
        End If

        
        '--- BLOQUE DE REVISIÓN OPCIONAL ---
        r = MsgBox("¿Desea revisar el registro en la tabla de Excel?", vbYesNo + vbQuestion)
        
        If r = vbYes Then
            ' Descargamos la tabla actualizada de Access a Excel
            Call DataBaseUtils.GetExcelTable(TSOLPEDs, refresh:=True)
            Dim camposVisibles As Variant
            
            camposVisibles = Array(CampoDB(TSOLPEDs, spd_solped), CampoDB(TSOLPEDs, spd_codigo), CampoDB(TSOLPEDs, spd_proveedor), _
                    CampoDB(TSOLPEDs, spd_proveedor_unico), CampoDB(TSOLPEDs, spd_capturo))
            Call StringUtils.SeleccionColumnasTabla(TSOLPEDs, camposVisibles, True)

        End If
        
        ' Cerramos el formulario solo si hubo éxito
        Unload Me
        Call Seguridad.LockSheet(ActiveSheet)
        App.EnfocarStockTechLabels
    Else
        ' Si AddSOLPED devuelve False sin error crítico de sistema
        MsgBox "El registro no pudo completarse. Verifique la conexión con la base de datos.", vbExclamation
    End If
    
    Exit Sub

'============================================================
' 3. Gestión de Errores y Trazabilidad
'============================================================
ErrorHandler:
    
    App.SystemError "ERROR REGISTRADO: " & Err.Description
    
End Sub
Private Sub UbicacionCodigo_Click()
    Call UbicacionCodigo_Change
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

Private Sub UserForm_Activate()
    GetMatrix
    
    Dim sh As Worksheet
    Set sh = ActiveSheet
    ' Antes de inicializarse esta interfáz ya se confirmó desde AccionesRibbon que es una Hoja autorizada y es TSolicitudes
    If Selection.Rows.Count = 1 Then ' Aunque sea una selección correcta se ejecutá el comando y dará 0
        If Cells(Selection.Row, 7) = False Then
            IdSolicitud = StringUtils.GetIdFromExcel(TSolicitudes)
            If IdSolicitud <> 0 Then
                TipoSolicitud = "Manual"
                
                Dim tipo As String
                tipo = Cells(Selection.Row, 5).Value
                Codigo = Cells(Selection.Row, 4).Value & ""
                Dim name As String
                If Codigo <> "Solicitar Código" Then
                    Me.RefactionOption.Enabled = False
                    Me.EquipOption.Enabled = False
                    Me.AreaOption.Enabled = False
                    Select Case LCase(tipo)
                        Case "equipo"
                            
                            Me.EquipOption.Value = True
                            
                            Call EquipOption_Click
                            
                            Me.ServicioEquipo.Visible = False
                            Me.ServicioEquipo.Enabled = False
                            
                            name = DataBaseUtils.DatoDataBase(TEquipos, CampoDB(TEquipos, equ_codigo), Codigo, CampoDB(TEquipos, equ_nombre))
                            If name <> "" Then
                                Me.EquipName.Value = name
                                Call EquipName_Change
                                Me.EquipName.Enabled = False
                            End If
                            
                        Case "refacción"
                            Me.RefactionOption.Value = True
                            
                            Call RefactionOption_Click
                            name = DataBaseUtils.DatoDataBase(TRefacciones, CampoDB(TRefacciones, ref_material), Codigo, CampoDB(TRefacciones, ref_descripcion))
                            If name <> "" Then
                                Me.RefaccionName.Value = name
                                Call RefaccionName_Click
                                Me.RefaccionName.Enabled = False
                            End If
                        Case "servicio a equipo"
                            Me.EquipOption.Value = True
                                                        
                            Call EquipOption_Change
                            Me.ServicioEquipo.Value = True
                            Me.ServicioEquipo.Enabled = False
                            name = DataBaseUtils.DatoDataBase(TEquipos, CampoDB(TEquipos, equ_codigo), Codigo, CampoDB(TEquipos, equ_nombre))
                            If name <> "" Then
                                Me.EquipName.Value = name
                                Call EquipName_Change
                                Me.EquipName.Enabled = False
                            End If
                        Case "servicio área"
                            Me.AreaOption.Value = True
                            
                            Call AreaOption_Click
                            name = DataBaseUtils.DatoDataBase(TUbicaciones, CampoDB(TUbicaciones, ubi_id_ubicacion), Codigo, CampoDB(TUbicaciones, ubi_ubicacion))
                            If name <> "" Then
                                Me.UbicacionCodigo.Value = name
                                Call UbicacionCodigo_Click
                                Me.UbicacionCodigo.Enabled = False
                            End If
                    End Select
                End If
                uso = Cells(Selection.Row, 9).Value & ""
                If uso <> "" Then
                    Me.UsoSOLPED.Text = "Mantenimiento " & uso
                End If
            End If
        End If
    End If
    
End Sub
Private Sub EquipName_Change()
    Dim fila As Long
    fila = EquipName.ListIndex
    
    ' =======================================================
    ' COINCIDENCIA EXACTA
    ' =======================================================
    ' Evaluamos si el texto escrito en la caja es estrictamente
    ' idéntico al elemento que VBA está intentando autoseleccionar.
    If fila <> -1 And Trim(EquipName.Value) <> "" Then
        
        ' UCase asegura que no falle por diferencias de mayúsculas/minúsculas
        If UCase(EquipName.Value) = UCase(EquipName.List(fila)) Then
            ' Solo si hay coincidencia perfecta, disparamos el motor
            Call RefreshEquipment
            Exit Sub
        End If
        
    End If
    
    ' =======================================================
    ' ESTADO DE ESPERA (Limpieza Visual)
    ' =======================================================
    ' Si el usuario sigue escribiendo o no hay coincidencia exacta,
    ' enviamos una bandera True para limpiar la pantalla y evitar fantasmas.
    Call RefreshEquipment(LimpiarPantalla:=True)
End Sub

' Mantenemos el Click como respaldo por si el operador usa el ratón
Private Sub EquipName_Click()
    Call EquipName_Change
End Sub
Function RefreshEquipment(Optional UltimoDato As Boolean = False, Optional LimpiarPantalla As Boolean = False)
    Dim fila As Long
    Dim Criticidad As String
    
    ' =======================================================
    ' MODO LIMPIEZA (El usuario está escribiendo o borró el dato)
    ' =======================================================
    If LimpiarPantalla Then
        Me.CodigoEquipo.Caption = ""
        MarcaEquipo.Caption = ""
        ModeloEquipo.Caption = ""
        Ubicacion.Caption = ""
        
        CriticidadRed.Visible = False
        CriticidadYellow.Visible = False
        CriticidadGreen.Visible = False
        Exit Function
    End If
    
    ' =======================================================
    ' MODO EXTRACCIÓN Y RENDERIZADO
    ' =======================================================
    fila = EquipName.ListIndex
    
    If UltimoDato Then
        fila = CInt(UBound(EquiposMatrix, 2))
    End If
    
    ' Inyección segura a los Labels
    Me.CodigoEquipo.Caption = EquiposMatrix(1, fila) & ""
    MarcaEquipo.Caption = EquiposMatrix(3, fila) & ""
    ModeloEquipo.Caption = EquiposMatrix(4, fila) & ""
    Ubicacion.Caption = EquiposMatrix(15, fila) & ""
    
    ' =======================================================
    ' SEMÁFORO DE CRITICIDAD
    ' =======================================================
    Criticidad = LCase(Trim(EquiposMatrix(8, fila) & ""))
    
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
            ' Asumimos que cualquier otro valor (o vacío) es criticidad baja/verde
            CriticidadRed.Visible = False
            CriticidadYellow.Visible = False
            CriticidadGreen.Visible = True
    End Select
    
End Function
Private Sub Numero_Click()
    fila = Numero.ListIndex
    
    If fila = -1 Then Exit Sub
    
    CountName.Caption = CuentasMatrix(5, fila) & ""
    Me.Saldo.Caption = "$ " & Format(ConsultasSQL.GetSaldoCuenta(Me.Numero.Text), "#,##0.00")
    
End Sub


Private Sub Numero_Change()
    Dim texto As Variant
    Dim Caracter As Variant
    Dim Largo As Integer
    On Error Resume Next
    texto = Me.Cuenta.Value
    Largo = Len(Me.Cuenta.Value)
    For i = 1 To Largo
        Caracter = Mid(texto, i, 1)
        If Caracter <> "" Then
            If Caracter < Chr(48) Or Caracter > Chr(57) Then
                Me.Cuenta.Value = Replace(texto, Caracter, "")
            End If
        End If
    Next i
    On Error GoTo 0
    Caracter = 0
    Caracter1 = 0
End Sub
Private Sub SOLPED_Change()
    Dim texto As Variant
    Dim Caracter As Variant
    Dim Largo As Integer
    On Error Resume Next
    texto = Me.SOLPED.Value
    Largo = Len(Me.SOLPED.Value)
    If Largo > 10 Then
        Me.SOLPED.Value = Left(Me.SOLPED.Value, 10)
        MsgBox "Verifique los 10 dígitos de la SOLPED", vbExclamation
    End If
    For i = 1 To Largo
        Caracter = Mid(texto, i, 1)
        If Caracter <> "" Then
            If Caracter < Chr(48) Or Caracter > Chr(57) Then
                Me.SOLPED.Value = Replace(texto, Caracter, "")
            End If
        End If
    Next i
    On Error GoTo 0
    Caracter = 0
    Caracter1 = 0
End Sub

Private Sub NuevoEquipo_Click()
    Unload Me
    NewEquip.Show
End Sub
Private Sub Presupuestos_Click()
    DataBaseUtils.GetExcelTable (TPresupuestos)
    Call Seguridad.LockSheet(ActiveSheet)
    Unload Me
    MsgBox "Ahora puede editar la tabla y actualizar la Base de Datos", vbInformation
End Sub
Private Sub ProveedorName_Change()
    fila = Me.ProveedorName.ListIndex
    If fila = -1 Then Exit Sub
    Me.ProveedorCodigo.Caption = ProveedorMatrix(1, fila) & ""
End Sub
Private Sub ProveedorName_Click()
    Call ProveedorName_Change
End Sub

Private Function getIdSolicitud() As Boolean
    Const PROC_NAME As String = "getIdSolicitud"
    
    Dim idInput As Variant ' Renombrado para evitar conflictos
    Dim campos As Variant
    Dim valores As Variant
    Dim sh As Worksheet
    Dim r As VbMsgBoxResult
    Dim idColumnDBName As String
    
    camposSolicitudes = CamposTablaDB(TSolicitudes)
    idColumnDBName = camposSolicitudes(sol_id_solicitud)
    
    On Error GoTo ErrorHandler

    ' Si IdSolicitud se conserva, ya se ha determinado que no hay una selección correcta
    If IdSolicitud = 0 Or IdSolicitud = Empty Then
        
        r = MsgBox("No hay una referencia de solicitud o la selección es inválida." & vbNewLine & _
                   "¿Desea colocarla manualmente? (Necesita revisar la lista de solicitudes)", _
                   vbExclamation + vbYesNo, "Validación StockTech")
        
        If r = vbYes Then
            
            ' ==========================================================
            ' RAMA A: SELECCIÓN MANUAL
            ' ==========================================================
            Seguridad.UnlockBook
            If ActiveSheet.name <> TSolicitudes Then
                
                If StringUtils.ExisteHoja(TSolicitudes) Then
                    Set sh = ActiveWorkbook.Worksheets(TSolicitudes)
                    Call DataBaseUtils.GetExcelTable(TSolicitudes, refresh:=True)
                Else
                    Call DataBaseUtils.GetExcelTable(TSolicitudes)
                End If
                Call Seguridad.LockSheet(ActiveSheet)
            End If
            Seguridad.LockBook
again:
            idInput = Application.InputBox( _
                Prompt:="Ingrese el renglón de la solicitud que quiere crear (Ver hoja " & UCase(TSolicitudes) & ")", _
                Title:="Búsqueda de ID", _
                Type:=1) ' Type 1 restringe a solo números
        
            If VarType(idInput) = vbBoolean And idInput = False Then
                MsgBox "Operación cancelada por el usuario.", vbInformation, "StockTech"
                IdSolicitud = Empty
                getIdSolicitud = True
                Exit Function
            Else
                
                Dim realizado As Boolean
                realizado = DataBaseUtils.DatoDataBase(TSolicitudes, idColumnDBName, Cells(idInput, 1).Value, "Realizada")
                If realizado = True Then
                    MsgBox "Referencia inválida. La solicitud indicada ya fue revisada", vbExclamation
                    GoTo again
                Else
                    Dim objetivoid As Integer
                    objetivoid = Cells(idInput, 1)
                    TipoSolicitud = DataBaseUtils.DatoDataBase(TSolicitudes, idColumnDBName, objetivoid, "Tipo")
                    Dim codigoRevision As String
                    codigoRevision = DataBaseUtils.DatoDataBase(TSolicitudes, idColumnDBName, objetivoid, "Código")
                    If TipoSolicitud = "EQUIPO" Then
                        If Me.CodigoEquipo <> codigoRevision Then
                            MsgBox "El código del EQUIPO no corresponde con la SOLICITUD ", vbExclamation
                            MsgBox "Revise el código del EQUIPO"
                            getIdSolicitud = True
                            Exit Function
                        End If
                    End If
                    If TipoSolicitud = "REFACCIÓN" Then
                        If Me.RefaccionCodigo <> codigoRevision Then
                            MsgBox "El código de la REFACCIÓN no corresponde con la SOLICITUD", vbExclamation
                            MsgBox "Revise el código de la REFACCÓN"
                            getIdSolicitud = True
                            Exit Function
                        End If
                    End If
                    If LCase(TipoSolicitud) = "servicio de área" Then
                        If Me.UbicacionCodigo <> codigoRevision Then
                            MsgBox "El código de la UBICACIÓN no corresponde con la SOLICITUD", vbExclamation
                            MsgBox "Revise el código de la UBICACIÓN"
                            getIdSolicitud = True
                            Exit Function
                        End If
                    End If
                End If
                
                IdSolicitud = DataBaseUtils.DatoDataBase(TSolicitudes, idColumnDBName, Cells(idInput, 1), idColumnDBName)
            End If
            
        Else
            TipoSolicitud = "Automática"
            ' ==========================================================
            ' RAMA B: CREACIÓN AUTOMÁTICA
            ' ==========================================================
            MsgBox "Esta acción generará una solicitud automática creada por " & UCase(Environ("USERNAME")), vbExclamation, "Registro Automático"
           
            campos = Array(camposSolicitudes(sol_fecha), camposSolicitudes(sol_usuario), camposSolicitudes(sol_codigo), _
                            camposSolicitudes(sol_tipo), camposSolicitudes(sol_unidades), camposSolicitudes(sol_realizada))
            valores = Array(Date, GetCurrentUser, "No aplica", "Automática", Me.Unidades.Value, True)
            
            
            ' Inserción utilizando Constantes
            
            If DataBaseUtils.AddRegister(TSolicitudes, campos, valores) Then
                
                ' Llamada a la función genérica
                IdSolicitud = DataBaseUtils.GetLastIdFromDB(TSolicitudes, idColumnDBName)
                ' Validación de integridad
                If IdSolicitud <= 0 Then
                    MsgBox "Error crítico: No se pudo recuperar el ID generado. Verifique la conexión.", vbCritical, "Fallo de Sistema"
                    DataBaseUtils.RollBack (TAuditTrail)
                    getIdSolicitud = False
                    Exit Function
                End If
                
            Else
                ' Si AddRegister falla (falla de red o SQL), levantamos un error para la Caja Negra
                Err.Raise vbObjectError + 515, , "Fallo en AddRegister al intentar crear solicitud automática."
            End If
        End If
    End If
    
    ' Si todo el proceso se ejecutó correctamente
    getIdSolicitud = True
    Exit Function
    
ErrorHandler:
    
    App.SystemError "ERROR REGISTRADO: " & Err.Description
    
    getIdSolicitud = False
End Function
Private Function AddSOLPED(ByVal IdSolicitud As Long) As Boolean
    Const PROC_NAME As String = "AddSOLPED"
    
    Dim campos As Variant
    Dim valores As Variant
    
    On Error GoTo ErrorHandler
    Dim camposSOLPEDs As Variant
    camposSOLPEDs = CamposTablaDB(TSOLPEDs)
    
    ' 1. Preparación de Arrays (Verifica el nombre de la columna SOLPED en Access)
    'Campos y valores tienen que cambiar con respecto al tipo de SOLPED
    Dim control As control
    For Each control In Me.Controls
        If TypeName(control) = "OptionButton" Then
            If control.Value Then
                Dim tipo As String
                tipo = UCase(CStr(control.Caption))
            End If
        End If
    Next
    campos = Array(camposSOLPEDs(spd_solped), camposSOLPEDs(spd_cuenta), camposSOLPEDs(spd_codigo), _
                    camposSOLPEDs(spd_proveedor), camposSOLPEDs(spd_proveedor_unico), camposSOLPEDs(spd_capturo), camposSOLPEDs(spd_id_solicitud), camposSOLPEDs(spd_uso), camposSOLPEDs(spd_tipo))
    Select Case tipo
        Case "EQUIPO"
            If Me.ServicioEquipo.Value = True Then
                tipo = "Servicio a Equipo"
            End If
            valores = Array(Me.SOLPED.Text, Me.Numero.Text, Me.CodigoEquipo.Caption, Me.ProveedorName.Value, Me.Unico.Value, GetCurrentUser, IdSolicitud, Me.UsoSOLPED.Text, tipo)
        Case "ÁREA"
            tipo = "Servicio de área"
            valores = Array(Me.SOLPED.Text, Me.Numero.Text, Me.CodigoUbic.Caption, Me.ProveedorName.Value, Me.Unico.Value, GetCurrentUser, IdSolicitud, Me.UsoSOLPED.Text, tipo)
        Case "REFACCIÓN"
            valores = Array(Me.SOLPED.Text, Me.Numero.Text, Me.RefaccionCodigo.Caption, Me.ProveedorName.Value, Me.Unico.Value, GetCurrentUser, IdSolicitud, Me.UsoSOLPED.Text & " para equipo con código: " & Me.CodigoEquipo, tipo)
        Case Else
            MsgBox "Error al añadir SOLPED", vbCritical
            AddSOLPED = False
            Exit Function
    End Select
        
    
    ' 2. Intento de Inserción
    If DataBaseUtils.AddRegister(TablasDB.TSOLPEDs, campos, valores) Then
        Dim camposSOLPEDTrack As Variant
        camposSOLPEDTrack = CamposTablaDB(TSOLPEDsTrack)
        Dim valoresTrack1 As Variant
        Dim valoresTrack2 As Variant
        Dim camposTrack As Variant
        camposTrack = Array(camposSOLPEDTrack(spt_usuario), camposSOLPEDTrack(spt_fecha), camposSOLPEDTrack(spt_estado), camposSOLPEDTrack(spt_solped))
        valoresTrack1 = Array(UCase(Environ("USERNAME")), Date, "CREADA", Me.SOLPED.Text)
        valoresTrack2 = Array(UCase(Environ("USERNAME")), Date, "POR APROBAR", Me.SOLPED.Text)
        Call DataBaseUtils.AddRegister(TSOLPEDsTrack, camposTrack, valoresTrack1)
        Call DataBaseUtils.AddRegister(TSOLPEDsTrack, camposTrack, valoresTrack2)
        
        If StringUtils.ExisteHoja(TSolicitudes) Then
            Dim sh As Worksheet
            Set sh = ActiveWorkbook.Worksheets(TSolicitudes)
            If Seguridad.GetStockTechSheetName(sh) = TSolicitudes Then
                DataBaseUtils.GetExcelTable TSolicitudes, refresh:=True
                Call Seguridad.LockSheet(ActiveSheet)
            End If
        End If
        
        If Not DataBaseUtils.SetDataByID(TSolicitudes, CampoDB(TSolicitudes, sol_realizada), True, IdSolicitud) Then
            DataBaseUtils.RollBack (TSOLPEDs)
            Exit Function
        End If
        AddSOLPED = True
        Exit Function
    Else
        ' 3. MANEJO DE FALLO Y ROLLBACK
        ' Registramos el fallo inicial en la Caja Negra
        App.SystemError "AddRegister rechazó la inserción."
        
        If TipoSolicitud = "Automática" Then
            If Not DataBaseUtils.RollBack(TSolicitudes) Then
                App.SystemError "Fallo Crítico de Rollback: El ID automático (" & IdSolicitud & ") quedó huérfano en la DB."
            End If
        End If
        
        AddSOLPED = False
        Exit Function
    End If

ErrorHandler:
    
    App.SystemError "ERROR REGISTRADO: " & Err.Description
    AddSOLPED = False
End Function



