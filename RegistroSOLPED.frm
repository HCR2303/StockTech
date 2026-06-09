VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegistroSOLPED 
   Caption         =   "Registro de Nueva SOLPED"
   ClientHeight    =   7848
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9588.001
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
Private IdSolicitud As Integer
Private TipoSolicitud As String
Const MODULE_NAME As String = "RegistroSOLPED FORM"

Function GetMatrix()
    Dim CuentasList() As Variant
    Dim EquiposList() As Variant
    Dim ProveedorList() As Variant
    Dim RefaccionList() As Variant
    Dim Filas As Long
    
    'Comandos para adquisición de tablas de datos
    Dim camposCuent As Variant
    Dim camposEqu As Variant
    Dim camposUbic As Variant
    camposCuent = CamposTablaDB(TCuentas)
    camposEqu = CamposTablaDB(TEquipos)
    camposUbic = CamposTablaDB(TUbicaciones)
    CuentasMatrix = DataBaseUtils.QueryTableConsult(TPresupuestos, TCuentas, camposCuent(cta_id_cuenta), camposCuent(cta_nombre))
    EquiposMatrix = DataBaseUtils.QueryTableConsult(TEquipos, TUbicaciones, camposUbic(ubi_id_ubicacion), camposUbic(ubi_ubicacion))
    ProveedorMatrix = DataBaseUtils.TableDataBase(TProveedores, "*")
    RefaccionMatrix = DataBaseUtils.TableDataBase(TRefacciones, "*")
    
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
        EquiposList(i) = EquiposMatrix(1, i) & ""
    Next
    CodigoEquipo.List = EquiposList
    
    'Comandos para adquisición de lista de PROVEEDORES
    Filas = UBound(ProveedorMatrix, 2)
    ReDim ProveedorList(0 To Filas)
    For i = 0 To Filas
        ProveedorList(i) = ProveedorMatrix(1, i) & ""
    Next
    ProveedorCodigo.List = ProveedorList
    
    'Comandos para adquisición de lista de REFACCIONES
    Filas = UBound(RefaccionMatrix, 2)
    ReDim RefaccionList(0 To Filas)
    For i = 0 To Filas
        RefaccionList(i) = RefaccionMatrix(1, i) & ""
    Next
    RefaccionCodigo.List = RefaccionList
    
End Function

Private Sub RefaccionCodigo_Change()
    Dim fila As Long
    fila = RefaccionCodigo.ListIndex
    
    ' =======================================================
    ' ESCUDO DE COINCIDENCIA EXACTA
    ' =======================================================
    ' Verificamos que el índice exista y que la caja no esté vacía.
    If fila <> -1 And Trim(RefaccionCodigo.Value) <> "" Then
        
        ' Comparamos el texto escrito con el texto real del elemento en la lista.
        If UCase(RefaccionCodigo.Value) = UCase(RefaccionCodigo.List(fila)) Then
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
Private Sub RefaccionCodigo_Click()
    Call RefaccionCodigo_Change
End Sub
Function RefreshRefaction(Optional UltimoDato As Boolean = False, Optional LimpiarPantalla As Boolean = False)
    Dim fila As Long
    Dim Criticidad As String
    
    ' =======================================================
    ' MODO LIMPIEZA (Señal de Espera)
    ' =======================================================
    If LimpiarPantalla Then
        RefaccionName.Caption = ""
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
    fila = RefaccionCodigo.ListIndex
    
    If UltimoDato Then
        fila = CInt(UBound(RefaccionMatrix, 2))
    End If
    
    ' Inyección segura a los Labels (previene Error 13 por nulos)
    RefaccionName.Caption = RefaccionMatrix(2, fila) & ""
    Stock.Caption = RefaccionMatrix(4, fila) & " " & RefaccionMatrix(3, fila)
    If UCase(RefaccionCodigo.Value) Like "SER" & "*" Then
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
Private Sub RefaccionNueva_Click()
    Unload Me
    NewRefaction.Show
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
                    ' ProveedorCodigo es opcional según lógica previa
                    If espacio.Name <> "ProveedorCodigo" And espacio.Visible = True Then
                        nuloCount = nuloCount + 1
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
            Err.Raise vbObjectError + 513, , "No se pudo generar el Id_Solicitud."
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
            
            Dim ws As Worksheet
            Dim lo As ListObject
            
            ' Usamos el nombre de la constante para evitar errores de texto
            Set ws = ActiveWorkbook.Worksheets(TSOLPEDs)
            
            On Error Resume Next
            Set lo = ws.ListObjects(1)
            On Error GoTo ErrorHandler
            
            If Not lo Is Nothing Then
                ' Posicionamiento y limpieza visual
                Call StringUtils.UltimoDatoTabla(lo, 0)
                Dim camposSOLPEDs As Variant
                camposSOLPEDs = CamposTablaDB(TSOLPEDs)
                ' Ocultamos columnas de control para el usuario final
                lo.ListColumns(camposSOLPEDs(sol_orden_compra)).Range.EntireColumn.Hidden = True
                lo.ListColumns(camposSOLPEDs(sol_no_factura)).Range.EntireColumn.Hidden = True
                lo.ListColumns(camposSOLPEDs(sol_costo)).Range.EntireColumn.Hidden = True
                lo.ListColumns(camposSOLPEDs(sol_id_solicitud)).Range.EntireColumn.Hidden = True
                
                MsgBox "Verifique su información y actualice la tabla de ser necesario.", vbInformation
            Else
                App.SystemError "La tabla [" & TablasDB.TSOLPEDs & "] no pudo ser instanciada en Excel."
            End If
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
                Codigo = Cells(Selection.Row, 4).Value
                If tipo = "EQUIPO" Then
                    If Codigo <> "Solicitar Código" Then
                        Me.CodigoEquipo.Value = Cells(Selection.Row, 4).Value
                        Me.CodigoEquipo.Enabled = False
                    End If
                End If
                If tipo = "REFACCIÓN" Then
                    If Codigo <> "Solicitar Código" Then
                        Me.RefaccionCodigo.Value = Cells(Selection.Row, 4).Value
                        Me.RefaccionCodigo.Enabled = False
                    End If
                End If
                If LCase(tipo) Like "serv" & "*" Then
                    Me.RefaccionCodigo.Value = Cells(Selection.Row, 4).Value
                    Me.RefaccionCodigo.Enabled = False
                    Me.Unidades.Enabled = False
                    Me.Unidades.Visible = False
                    Me.UnidadesLabel.Visible = False
                End If
            End If
        End If
    End If
    
End Sub
Private Sub CodigoEquipo_Change()
    Dim fila As Long
    fila = CodigoEquipo.ListIndex
    
    ' =======================================================
    ' COINCIDENCIA EXACTA
    ' =======================================================
    ' Evaluamos si el texto escrito en la caja es estrictamente
    ' idéntico al elemento que VBA está intentando autoseleccionar.
    If fila <> -1 And Trim(CodigoEquipo.Value) <> "" Then
        
        ' UCase asegura que no falle por diferencias de mayúsculas/minúsculas
        If UCase(CodigoEquipo.Value) = UCase(CodigoEquipo.List(fila)) Then
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
Private Sub CodigoEquipo_Click()
    Call CodigoEquipo_Change
End Sub
Function RefreshEquipment(Optional UltimoDato As Boolean = False, Optional LimpiarPantalla As Boolean = False)
    Dim fila As Long
    Dim Criticidad As String
    
    ' =======================================================
    ' MODO LIMPIEZA (El usuario está escribiendo o borró el dato)
    ' =======================================================
    If LimpiarPantalla Then
        EquipName.Caption = ""
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
    fila = CodigoEquipo.ListIndex
    
    If UltimoDato Then
        fila = CInt(UBound(EquiposMatrix, 2))
    End If
    
    ' Inyección segura a los Labels
    EquipName.Caption = EquiposMatrix(2, fila) & ""
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
    Saldo.Caption = "$ " & ConsultasSQL.GetSaldoCuenta(Numero.Text) & ""
    
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
Private Sub ProveedorCodigo_Click()
    fila = ProveedorCodigo.ListIndex
    If fila = -1 Then Exit Sub
    ProveedorName.Caption = ProveedorMatrix(2, fila) & ""
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
            If ActiveSheet.Name <> TSolicitudes Then
                If StringUtils.ExisteHoja(TSolicitudes) Then
                    Set sh = ActiveWorkbook.Worksheets(TSolicitudes)
                    If Seguridad.GetStockTechSheetName(sh) = TSolicitudes Then
                        Call DataBaseUtils.GetExcelTable(TSolicitudes, refresh:=True)
                    Else
                        Call DataBaseUtils.GetExcelTable(TSolicitudes)
                    End If
                Else
                    Call DataBaseUtils.GetExcelTable(TSolicitudes)
                End If
                Call Seguridad.LockSheet(ActiveSheet)
            End If
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
                realizado = DataBaseUtils.DatoDataBase(TSolicitudes, idColumnDBName, Cells(idInput, 1), "Realizada")
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
    campos = Array(camposSOLPEDs(spd_solped), camposSOLPEDs(spd_cuenta), camposSOLPEDs(spd_codigo_equipo), camposSOLPEDs(SPD_REFACCION), _
                    camposSOLPEDs(spd_proveedor), camposSOLPEDs(spd_proveedor_unico), camposSOLPEDs(spd_capturo), camposSOLPEDs(spd_id_solicitud), camposSOLPEDs(spd_uso))
    valores = Array(Me.SOLPED.Text, Me.Numero.Text, Me.CodigoEquipo.Text, Me.RefaccionCodigo.Text, Me.ProveedorName.Caption, Me.Unico.Value, GetCurrentUser, IdSolicitud, Me.UsoSOLPED.Text)
    
    ' 2. Intento de Inserción
    If DataBaseUtils.AddRegister(TablasDB.TSOLPEDs, campos, valores) Then
        Dim camposSOLPEDTrack As Variant
        camposSOLPEDTrack = CamposTablaDB(TSOLPEDsTrack)
        Dim valoresTrack1 As Variant
        Dim valoresTrack2 As Variant
        Dim camposTrack As Variant
        camposTrack = Array(camposSOLPEDTrack(spt_usuario), camposSOLPEDTrack(spt_fecha), camposSOLPEDTrack(spt_estado), camposSOLPEDTrack(spt_solped))
        valoresTrack1 = Array(UCase(Environ("USERNAME")), StringUtils.EstablecerFecha(), "CREADA", Me.SOLPED.Text)
        valoresTrack2 = Array(UCase(Environ("USERNAME")), Date, "POR APROBAR", Me.SOLPED.Text)
        Call DataBaseUtils.AddRegister(TSOLPEDsTrack, camposTrack, valoresTrack1)
        Call DataBaseUtils.AddRegister(TSOLPEDsTrack, camposTrack, valoresTrack2)
        If TipoSolicitud <> "Automática" Then
        
            If DataBaseUtils.SetDataByID(TablasDB.TSolicitudes, CamposTablaDB(TSolicitudes)(sol_realizada), True, IdSolicitud) = False Then
                DataBaseUtils.RollBack (TSOLPEDs)
                
                Exit Function
            End If
            If StringUtils.ExisteHoja(TSolicitudes) Then
                Dim sh As Worksheet
                Set sh = ActiveWorkbook.Worksheets(TSolicitudes)
                If Seguridad.GetStockTechSheetName(sh) = TSolicitudes Then
                    Call DataBaseUtils.GetExcelTable(TSolicitudes, refresh:=True)
                    Call Seguridad.LockSheet(ActiveSheet)
                End If
            End If
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



