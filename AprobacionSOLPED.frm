VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AprobacionSOLPED 
   Caption         =   "Aprobación de SOLPED's"
   ClientHeight    =   7308
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10824
   OleObjectBlob   =   "AprobacionSOLPED.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "AprobacionSOLPED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CuentasMatrix As Variant
Private EquiposMatrix As Variant
Private ProveedorMatrix As Variant
Private RefaccionMatrix As Variant
Private SOLPEDMatrix As Variant

Const MODULE_NAME As String = "AprobacionSOLPED Frm"

Function GetMatrix()
    Dim SOLPEDMatrix As Variant
    Dim sqlQuery As String
    Dim r As VbMsgBoxResult
    
    ' =======================================================
    ' 1. ENSAMBLAJE DE CONSULTA (Blindaje Total contra Nulos)
    ' =======================================================
    ' Se añade el escudo "(... OR ... IS NULL)" al campo de aprobación
    ' para garantizar que las SOLPEDs recién nacidas no sean cortadas de la lista.
    sqlQuery = "SELECT DISTINCT [" & CampoDB(TSOLPEDs, spd_solped) & "] " & _
               "FROM [" & TSOLPEDs & "] " & _
               "WHERE ([" & CampoDB(TSOLPEDs, spd_aprobacion) & "] = 0 OR [" & CampoDB(TSOLPEDs, spd_aprobacion) & "] IS NULL) " & _
               "AND ([" & CampoDB(TSOLPEDs, spd_comentario) & "] NOT LIKE 'CANCELADA%' " & _
               "OR [" & CampoDB(TSOLPEDs, spd_comentario) & "] IS NULL)"
               
    ' =======================================================
    ' 2. EXTRACCIÓN Y RENDERIZADO
    ' =======================================================
    ' Disparamos la extracción con el motor ADO seguro
    SOLPEDMatrix = GetFromSQL(sqlQuery)
    
    ' =======================================================
    ' 3. VALIDACIÓN DE VACÍO (Control de Flujo)
    ' =======================================================
    If IsEmpty(SOLPEDMatrix) Then
        MsgBox "No hay SOLPED's por aprobar", vbInformation, "Estado de Operación"
        r = MsgBox("¿Desea ver la tabla de SOLPED's?", vbYesNo + vbQuestion, "Sin Registros Pendientes")
        
        If r = vbYes Then
            Unload Me
            DataBaseUtils.GetExcelTable TSOLPEDs, refresh:=True
            
            Dim camposNoVis As Variant
            camposNoVis = Array(spd_id_solicitud, spd_comentario, spd_aprobacion, spd_comentario)
            Call StringUtils.OcultarColumnasTabla(TSOLPEDs, camposNoVis)
            Call Seguridad.LockSheet(ActiveSheet)
        Else
            Unload Me
        End If
        
        Exit Function
    End If
    
    ' =======================================================
    ' 4. INYECCIÓN A LA INTERFAZ (Prevención de Corte Visual)
    ' =======================================================
    ' Apagamos estrictamente la propiedad ColumnHeads por código.
    ' Esto asegura que VBA no consuma el primer registro de la matriz
    ' utilizándolo como un título invisible.
    Me.SOLPED.ColumnHeads = False
    
    ' La propiedad .Column recibe nativamente la matriz (Cols, Filas) de ADO
    Me.SOLPED.Column = SOLPEDMatrix

End Function
Private Sub Aprobacion_Click()
    Dim spd As String
    spd = CStr(Trim(Me.SOLPED.Text))
    If spd = "" Then
        MsgBox "Se debe seleccionar una SOLPED primero", vbExclamation
        Exit Sub
    End If
    If DataBaseUtils.LogDB("Aprobación", "SOLPED: " & spd, "Se aprueba la SOLPED " & spd) Then
        If DataBaseUtils.SetRegister(TSOLPEDs, CampoDB(TSOLPEDs, spd_aprobacion), True, CampoDB(TSOLPEDs, spd_solped), spd) Then
            Dim campos As Variant
            Dim valores As Variant
            campos = CamposTablaDB(TSOLPEDsTrack)
            StringUtils.DropFromArray campos, 0
            StringUtils.DropFromArray campos, UBound(campos)
            valores = Array(UCase(Environ("USERNAME")), Date, "APROBADA", spd)
            If DataBaseUtils.AddRegister(TSOLPEDsTrack, campos, valores) Then
                MsgBox "SOLPED: " & spd & " aprobada exitosamente"
                Unload Me
                AprobacionSOLPED.Show
            Else
                Call DataBaseUtils.SetRegister(TSOLPEDs, CampoDB(TSOLPEDs, spd_aprobacion), False, CampoDB(TSOLPEDs, spd_solped), spd)
                DataBaseUtils.RollBack (TAuditTrail)
            End If
        End If
    Else
        App.SystemError ("Error al registrar AuditTrail")
    End If
    
End Sub

Private Sub Cancelar_Click()
    Dim spd As String
    spd = CStr(Trim(Me.SOLPED.Text))
    If spd = "" Then
        MsgBox "Se debe seleccionar una SOLPED primero", vbExclamation
        Exit Sub
    End If
    Dim Comentario As String
    With AuditTrail
        .Caption = "CANCELACIÓN SOLPED"
        .AccionLabel = "Cancelación de SOLPED: " & spd
        .Resultado = ""
        .Show
        Comentario = .Resultado
    End With
    If Trim(Comentario) = "" Then
        MsgBox "Acción cancelada", vbInformation
        Exit Sub
    Else
        If DataBaseUtils.LogDB("Cancelación", "SOLPED: " & spd, "Se cancela la SOLPED: " & spd & ", Razón: " & Comentario) Then
        If DataBaseUtils.SetRegister(TSOLPEDs, CampoDB(TSOLPEDs, spd_comentario), "CANCELADA por Gerencia, Razon: " & Comentario, CampoDB(TSOLPEDs, spd_solped), spd) Then
            Dim campos As Variant
            Dim valores As Variant
            campos = CamposTablaDB(TSOLPEDsTrack)
            StringUtils.DropFromArray campos, 0
            StringUtils.DropFromArray campos, UBound(campos)
            valores = Array(UCase(Environ("USERNAME")), Date, "CANCELADA", spd)
            If DataBaseUtils.AddRegister(TSOLPEDsTrack, campos, valores) Then
                MsgBox "SOLPED: " & spd & " cancelada exitosamente"
                Unload Me
                AprobacionSOLPED.Show
            Else
                Call DataBaseUtils.SetRegister(TSOLPEDs, CampoDB(TSOLPEDs, spd_comentario), "", CampoDB(TSOLPEDs, spd_solped), spd)
                DataBaseUtils.RollBack (TAuditTrail)
            End If
        Else
            DataBaseUtils.RollBack (TAuditTrail)
        End If
    Else
        App.SystemError ("Error al registrar AuditTrail")
    End If
    End If
    
End Sub

Private Sub RefaccionesCods_Click()
    Call RefaccionesCods_Change
End Sub
Private Sub RefaccionesCods_Change()
    Dim arrSeleccionadas As Variant
    Dim totalSeleccionadas As Long
    Dim camposRef As Variant
    Dim busqueda As String
    Dim datosRefaccion As Variant
    Dim Criticidad As String
    
    camposRef = TablasDB.CamposTablaDB(TRefacciones)
    
    ' 1. Consultamos cuántas filas están marcadas actualmente
    arrSeleccionadas = RefaccionesSeleccionadas()
    
    ' 2. Salida rápida si la matriz está vacía (evita errores en cadena)
    If IsEmpty(arrSeleccionadas) Then
        ' Limpieza visual si el usuario desmarca todo
        Me.Ramarrillo.Visible = False
        Me.Rrojo.Visible = False
        Me.Rverde.Visible = False
        Me.Nulo.Visible = False
        Me.RefaccionName.Visible = False
        Me.Stock.Visible = False
        Me.Rubicacion.Visible = False
        
        
        
        Exit Sub
    End If
    
    totalSeleccionadas = UBound(arrSeleccionadas) + 1
    
    ' 3. Evaluación del estado de la interfaz
    If totalSeleccionadas = 1 Then
    
        ' CORRECCIÓN 1: Extraemos del arreglo, no de la propiedad .Value
        busqueda = arrSeleccionadas(0)
        
        ' CORRECCIÓN 5: Añadidas las comillas simples para el parámetro SQL
        datosRefaccion = DataBaseUtils.TableDataBase(TRefacciones, "*", False, camposRef(ref_material) & " = '" & busqueda & "'")
        
        If Not IsEmpty(datosRefaccion) Then
            ' CORRECCIÓN 4: Restauramos la visibilidad de los controles
            Me.RefaccionName.Visible = True
            Me.Stock.Visible = True
            Me.Rubicacion.Visible = True
            
            ' CORRECCIÓN 2: Uso correcto del Enum con índice bidimensional (Columna, Fila 0)
            On Error Resume Next
            Me.RefaccionName.Caption = datosRefaccion(CNameRefacciones.ref_descripcion, 0)
            Me.Stock.Caption = datosRefaccion(CNameRefacciones.ref_stock, 0)
            Me.Rubicacion.Caption = datosRefaccion(CNameRefacciones.ref_ubicacion_almacen, 0)
            
            
            Criticidad = LCase(datosRefaccion(CNameRefacciones.ref_criticidad, 0) & "")
            On Error GoTo 0
            ' 1. Limpieza de estado visual (Apagamos todo por defecto)
            Me.Rrojo.Visible = False
            Me.Ramarrillo.Visible = False
            Me.Rverde.Visible = False
            Me.Nulo.Visible = False
            
            ' 2. Evaluación estricta
            Select Case Criticidad
                Case "alto"
                    Me.Rrojo.Visible = True
                Case "medio"
                    Me.Ramarrillo.Visible = True
                Case "bajo" ' Asegúrate de usar la palabra exacta de tu DB para el verde
                    Me.Rverde.Visible = True
                Case Else
                    Me.Nulo.Visible = True
            End Select
        End If
        
    Else
        ' Múltiples seleccionadas: Ocultamos los datos específicos
        Me.Ramarrillo.Visible = False
        Me.Rrojo.Visible = False
        Me.Rverde.Visible = False
        Me.Nulo.Visible = False
        Me.RefaccionName.Visible = False
        Me.Stock.Visible = False
        Me.Rubicacion.Visible = False
        
        
        
    End If

End Sub
Private Sub SOLPED_Change()
    
    fila = SOLPED.ListIndex
    
    If fila <> -1 Or SOLPED <> Empty Then
        CargarDatosSOLPED (SOLPED.Value)
    End If
End Sub

Private Sub UserForm_Activate()
    Me.Aprobacion.BackColor = RGB(204, 255, 204)
    Me.Cancelar.BackColor = RGB(255, 200, 200)
    GetMatrix
End Sub
Public Sub CargarDatosSOLPED(ByVal idSolpedSeleccionada As String)
    Dim arrSpd As Variant
    Dim arrEqu As Variant
    Dim arrCta As Variant
    Dim arrProv As Variant
    Dim camposUbic As Variant
    
    ' =======================================================
    ' PASO 1: EXTRACCIÓN MAESTRA
    ' =======================================================
    arrSpd = DataBaseUtils.TableDataBase(TSOLPEDs, "*", False, "[" & CampoDB(TSOLPEDs, spd_solped) & "] = '" & idSolpedSeleccionada & "'")
    
    If IsEmpty(arrSpd) Then Exit Sub
    
    ' =======================================================
    ' PASO 2: LIMPIEZA VISUAL DEL ENTORNO
    ' =======================================================
    Me.Refacciones.Visible = False
    Me.Equipo.Visible = False
    Me.Area.Visible = False
    
    ' =======================================================
    ' PASO 3: EXTRACCIÓN SEGURA DE MATRICES (Blindaje de Índices)
    ' =======================================================
    Dim valTipo As Variant
    Dim valCodigo As Variant
    Dim valProveedor As Variant
    Dim valUnico As Variant
    Dim valCuenta As Variant
    
    Dim tipoSolped As String
    Dim Codigo As String
    
    ' Obtenemos las matrices de la consulta principal
    valTipo = DataBaseUtils.UniqueValues(spd_tipo, arrSpd)
    valCodigo = DataBaseUtils.UniqueValues(spd_codigo, arrSpd)
    valProveedor = DataBaseUtils.UniqueValues(CNameSOLPEDs.spd_proveedor, arrSpd)
    valUnico = DataBaseUtils.UniqueValues(CNameSOLPEDs.spd_proveedor_unico, arrSpd)
    valCuenta = DataBaseUtils.UniqueValues(CNameSOLPEDs.spd_cuenta, arrSpd)
    
    ' Validamos y casteamos el tipo a texto plano sacando el índice (0)
    If Not IsEmpty(valTipo) Then tipoSolped = LCase(Trim(valTipo(0) & ""))
    If Not IsEmpty(valCodigo) Then Codigo = CStr(valCodigo(0) & "")
    
    ' =======================================================
    ' PASO 4: MOTOR DE RUTEO (Mutación de Interfaz)
    ' =======================================================
    Select Case tipoSolped
        
        ' ---------------------------------------------------
        ' RUTA A: EQUIPOS Y SERVICIOS
        ' ---------------------------------------------------
        Case "servicio a equipo", "equipo"
            Me.Equipo.Visible = True
            ' Ahora sí asignamos el código del equipo a la interfaz
            Me.CodigoEquipo.Caption = Codigo
            
            arrEqu = DataBaseUtils.TableDataBase(TEquipos, "*", False, CampoDB(TEquipos, equ_codigo) & " = '" & Codigo & "'")
            
            If Not IsEmpty(arrEqu) Then
                ' Inyección segura con (& "")
                Me.MarcaEquipo.Caption = arrEqu(CNameEquipos.equ_marca, 0) & ""
                Me.ModeloEquipo.Caption = arrEqu(CNameEquipos.equ_modelo, 0) & ""
                Me.EquipName.Caption = arrEqu(CNameEquipos.equ_nombre, 0) & ""
                
                Dim Criticidad As String
                Criticidad = LCase(Trim(arrEqu(CNameEquipos.equ_criticidad, 0) & ""))
                
                ' Semáforo de Criticidad
                CriticidadRed.Visible = False
                CriticidadYellow.Visible = False
                CriticidadGreen.Visible = False
                
                Select Case Criticidad
                    Case "alto": CriticidadRed.Visible = True
                    Case "medio": CriticidadYellow.Visible = True
                    Case Else: CriticidadGreen.Visible = True
                End Select
                
                ' Extracción de Ubicación
                camposUbic = TablasDB.CamposTablaDB(TUbicaciones)
                Dim ubic As String
                ' Aseguramos el uso correcto de las fábricas sin mezclar tipos
                ubic = DataBaseUtils.DatoDataBase(TUbicaciones, camposUbic(ubi_id_ubicacion), arrEqu(CNameEquipos.equ_id_ubicacion, 0), camposUbic(ubi_ubicacion))
                Me.Ubicacion.Caption = ubic & ""
            End If
            
        ' ---------------------------------------------------
        ' RUTA B: REFACCIONES
        ' ---------------------------------------------------
        Case "refacción"
            Me.Refacciones.Visible = True
            Me.RefaccionesCods.Clear
            
            ' valCodigo ya contiene la matriz de refacciones de esta SOLPED
            If Not IsEmpty(valCodigo) Then
                If UBound(valCodigo) = 0 Then
                    Me.RefaccionesCods.AddItem valCodigo(0)
                Else
                    Me.RefaccionesCods.List = valCodigo
                End If
            End If
            
        ' ---------------------------------------------------
        ' RUTA C: SERVICIO DE ÁREA
        ' ---------------------------------------------------
        Case "servicio de área"
            Me.Area.Visible = True
            
            camposUbic = TablasDB.CamposTablaDB(TUbicaciones)
            ' Corrección de errores tipográficos en las fábricas
            ubic = DataBaseUtils.DatoDataBase(TUbicaciones, camposUbic(ubi_id_ubicacion), Codigo, camposUbic(ubi_ubicacion))
            
            Me.NombreUbic.Caption = ubic & ""
            Me.CodigoUbic.Caption = Codigo
    End Select
        
    ' =======================================================
    ' PASO 5: RENDERIZADO DE DATOS GLOBALES
    ' =======================================================
    If Not IsEmpty(valCuenta) Then
        Me.NumeroCuenta.Caption = valCuenta(0) & ""
        arrCta = TablasDB.CamposTablaDB(TCuentas)
        Me.CountName.Caption = DataBaseUtils.DatoDataBase(TCuentas, arrCta(cta_id_cuenta), valCuenta(0), arrCta(cta_nombre)) & ""
    End If
    
    ' Si el proveedor único es relevante para tu interfaz, lo asignarías aquí:
    ' If Not IsEmpty(valProveedor) Then Me.ProveedorLabel.Caption = valProveedor(0) & ""
       
End Sub
Public Function RefaccionesSeleccionadas() As Variant
    Dim i As Long
    Dim conteo As Long
    Dim arrSeleccionados() As String
    
    If Me.RefaccionesCods.ListCount = 0 Then
        RefaccionesSeleccionadas = Empty
        Exit Function
    End If
    ' 1. Primer recorrido: Determinar la longitud del arreglo
    conteo = 0
    ' Nota: ListCount devuelve el total, pero el índice empieza en 0
    For i = 0 To Me.RefaccionesCods.ListCount - 1
        If Me.RefaccionesCods.Selected(i) Then
            conteo = conteo + 1
        End If
    Next i
    
    ' 2. Salida rápida si no hay selecciones
    If conteo = 0 Then
        RefaccionesSeleccionadas = Empty
        Exit Function
    End If
    
    ' 3. Dimensionamiento estricto de la memoria
    ReDim arrSeleccionados(0 To conteo - 1)
    
    ' 4. Segundo recorrido: Extracción de los datos
    Dim indiceArreglo As Long
    indiceArreglo = 0
    
    For i = 0 To Me.RefaccionesCods.ListCount - 1
        If Me.RefaccionesCods.Selected(i) Then
            ' Capturamos el texto de la fila seleccionada
            arrSeleccionados(indiceArreglo) = Me.RefaccionesCods.List(i)
            indiceArreglo = indiceArreglo + 1
        End If
    Next i
    
    RefaccionesSeleccionadas = arrSeleccionados
End Function
