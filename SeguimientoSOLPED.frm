VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SeguimientoSOLPED 
   Caption         =   "Registro de SOLPED"
   ClientHeight    =   12396
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10992
   OleObjectBlob   =   "SeguimientoSOLPED.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "SeguimientoSOLPED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CuentasMatrix As Variant
Private EquiposMatrix As Variant
Private ProveedorMatrix As Variant
Private RefaccionMatrix As Variant
Private SOLPEDMatrix As Variant

Const MODULE_NAME As String = "SeguimientoSOLPED FORM"

Function GetMatrix(Optional llegada As Boolean = False, Optional ByVal contratiempo As Boolean = False)
    Dim SOLPEDMatrix As Variant
    Dim sqlQuery As String
    
    ' 1. Construimos el SQL dinámico
    sqlQuery = "SELECT DISTINCT [" & CampoDB(TSOLPEDs, spd_solped) & "] " & _
               "FROM [" & TSOLPEDs & "] " & _
               "WHERE [" & CampoDB(TSOLPEDs, spd_aprobacion) & "] = -1 AND " & _
               CampoDB(TSOLPEDs, spd_costo) & " IS NULL"
    If llegada Then
        sqlQuery = "SELECT DISTINCT [" & CampoDB(TSOLPEDs, spd_solped) & "] " & _
               "FROM [" & TSOLPEDs & "] " & _
               "WHERE [" & CampoDB(TSOLPEDs, spd_llegada) & "] = 0 AND " & _
               CampoDB(TSOLPEDs, spd_costo) & " IS NOT NULL"
    End If
    If contratiempo Then
        sqlQuery = "SELECT DISTINCT [" & CampoDB(TSOLPEDs, spd_solped) & "] " & _
               "FROM [" & TSOLPEDs & "] " & _
               "WHERE [" & CampoDB(TSOLPEDs, spd_llegada) & "] = 0 "
    End If
    ' 2. Disparamos la extracción
    SOLPEDMatrix = GetFromSQL(sqlQuery)
    
    ' 3. Validación de vacío
    If IsEmpty(SOLPEDMatrix) Then
        MsgBox "No hay SOLPED's aprobadas", vbInformation
        r = MsgBox("¿Desea ver las SOLPED's pendientes por aprobar?", vbYesNo, "Sin Registros Pendientes")
        If r = vbYes Then
            Unload Me
            DataBaseUtils.GetExcelTable TSOLPEDs, refresh:=True
            Call StringUtils.FiltrarTabla(TSOLPEDs, CampoDB(TSOLPEDs, spd_aprobacion), False)
            Dim camposNoVis As Variant
            camposNoVis = Array(spd_id_solicitud, spd_comentario, spd_costo, spd_no_factura, spd_orden_compra)
            Call StringUtils.OcultarColumnasTabla(TSOLPEDs, camposNoVis)
            Call Seguridad.LockSheet(ActiveSheet)
        End If
        Exit Function
    End If
    
    ' La propiedad .Column recibe nativamente la matriz (Cols, Filas) de ADO
    Me.SOLPED.Column = SOLPEDMatrix

End Function
Private Sub ConOC_Click()
    Call GetMatrix
    If Me.ConOC.Value Then
        Me.ObservacionLabel.Visible = False
        Me.observaciones.Visible = False
    End If
End Sub
Private Sub llegada_Click()
    If Me.llegada.Value Then
        Call GetMatrix(True)
        Me.ObservacionLabel.Visible = False
        Me.observaciones.Visible = False
    End If
End Sub

Private Sub Proceso_Click()
    Call GetMatrix(contratiempo:=Proceso.Value)
    If Proceso.Value Then
        Me.ObservacionLabel.Visible = True
        Me.observaciones.Visible = True
    End If
End Sub
Private Sub RefaccionesCods_Change()
    Dim arrSeleccionadas As Variant
    Dim totalSeleccionadas As Long
    Dim camposRef As Variant
    Dim busqueda As String
    Dim datosRefaccion As Variant
    Dim Criticidad As String
    
    camposRef = TablasDB.CamposTablaDB(TRefacciones)
    
    ' =======================================================
    ' 1. ESCANEO DE INTERFAZ
    ' =======================================================
    ' Consultamos cuántas filas están marcadas actualmente
    arrSeleccionadas = RefaccionesSeleccionadas()
    
    ' Salida rápida si la matriz está vacía (evita errores en cadena)
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
    
    ' =======================================================
    ' 2. EVALUACIÓN DE SELECCIÓN ÚNICA
    ' =======================================================
    If totalSeleccionadas = 1 Then
    
        ' Extraemos del arreglo, no de la propiedad .Value
        busqueda = arrSeleccionadas(0)
        
        ' Ejecutamos la consulta con el parámetro SQL protegido
        datosRefaccion = DataBaseUtils.TableDataBase(TRefacciones, "*", False, camposRef(ref_material) & " = '" & busqueda & "'")
        
        ' --- CONDICIONAL PRINCIPAL ---
        If Not IsEmpty(datosRefaccion) Then
            
            ' A. EXITO EN BÚSQUEDA: Restauramos la visibilidad de los controles
            Me.RefaccionName.Visible = True
            Me.Stock.Visible = True
            Me.Rubicacion.Visible = True
            
            ' Inyección segura de datos concatenando ("") para evitar Error 13 por Nulls
            Me.RefaccionName.Caption = datosRefaccion(CNameRefacciones.ref_descripcion, 0) & ""
            Me.Stock.Caption = datosRefaccion(CNameRefacciones.ref_stock, 0) & ""
            Me.Rubicacion.Caption = datosRefaccion(CNameRefacciones.ref_ubicacion_almacen, 0) & ""
            
            Criticidad = LCase(Trim(datosRefaccion(CNameRefacciones.ref_criticidad, 0) & ""))
            
            ' Motor de semáforo visual
            Me.Rrojo.Visible = False
            Me.Ramarrillo.Visible = False
            Me.Rverde.Visible = False
            Me.Nulo.Visible = False
            
            Select Case Criticidad
                Case "alto"
                    Me.Rrojo.Visible = True
                Case "medio"
                    Me.Ramarrillo.Visible = True
                Case "bajo"
                    Me.Rverde.Visible = True
                Case Else
                    Me.Nulo.Visible = True
            End Select
        
        Else
            ' =======================================================
            ' 3. BLINDAJE CONTRA INFORMACIÓN FANTASMA (CORRECCIÓN CRÍTICA)
            ' =======================================================
            ' B. FALLO EN BÚSQUEDA: El código no existe en la base de datos.
            ' Ocultamos los elementos para que el usuario no vea la refacción del clic anterior.
            Me.Ramarrillo.Visible = False
            Me.Rrojo.Visible = False
            Me.Rverde.Visible = False
            Me.Nulo.Visible = False
            Me.RefaccionName.Visible = False
            Me.Stock.Visible = False
            Me.Rubicacion.Visible = False
        End If
        
    Else
        ' =======================================================
        ' 4. EVALUACIÓN DE SELECCIÓN MÚLTIPLE
        ' =======================================================
        ' Múltiples seleccionadas: Ocultamos los datos específicos
        Me.Ramarrillo.Visible = False
        Me.Rrojo.Visible = False
        Me.Rverde.Visible = False
        Me.Nulo.Visible = False
        Me.RefaccionName.Visible = False
        Me.Stock.Visible = False
        Me.Rubicacion.Visible = False
        
        If Me.ConOC.Value Then
            Me.ObservacionLabel.Visible = False
            Me.observaciones.Visible = False
        End If
        
    End If

End Sub

Private Sub RefaccionesCods_Click()
    Call RefaccionesCods_Change
End Sub
Private Sub RegistroSeguimiento_Click()
    Dim ctr As control
    Dim c As Long
    c = 0
    For Each ctr In Me.Controls
        If TypeName(ctr) = "OptionButton" Then
            If ctr.Value Then
                c = c + 1
            End If
        End If
    Next
    If c <> 1 Then
        MsgBox "Debe seleccionar y llenar los campos de Estado de SOLPED", vbExclamation, "Estado de SOLPED's"
        Exit Sub
    End If
    
    Dim Refacciones As Variant
        
    If Me.SOLPED.Text = Empty Then
        MsgBox "Debe seleccionar una SOLPED", vbExclamation
        Exit Sub
    End If
    spd = Me.SOLPED.Text
    Dim tabla As String
    tabla = TSOLPEDsTrack
    Dim registroExitoso As Boolean
    Dim con As control
    For Each con In Me.Controls
        If TypeName(con) = "Frame" And con.Visible = True Then
            If Me.Proceso.Value Then
                Call SOLPEDProceso
                Exit Sub
            End If
            If Me.llegada.Value Then
                Call SOLPEDLlegada
                Exit Sub
            End If
            Select Case con.name
                Case "Area"
                
                    With CostoRefaccion
                        .SOLPEDLabel.Caption = .SOLPEDLabel.Caption & Me.SOLPED.Text
                        .TipoLabel.Caption = "Área: " & Me.CodigoUbic.Caption
                        .Show
                        registroExitoso = .registroSOLPED
                    End With
                    
                    If registroExitoso Then
                        
                        campos = Array(CampoDB(tabla, spt_usuario), CampoDB(tabla, spt_fecha), CampoDB(tabla, spt_estado), CampoDB(tabla, spt_solped), CampoDB(tabla, spt_comentario))
                        ' Usamos la variable obsText que rescatamos en el Paso 1
                        valores = Array(GetCurrentUser(), StringUtils.EstablecerFecha("Generación de OC"), "CON OC", spd, "La generación de la OC es en la fecha de registro")
                        
                        If DataBaseUtils.LogDB("Cierre de SOLPED", "SOLPED: " & spd, "Se registra estado de CON OC de SOLPED: " & spd) Then
                            If DataBaseUtils.AddRegister(tabla, campos, valores) Then
                                r = MsgBox("¿Desea registrar una FECHA ESTIMADA DE LLEGADA?", vbYesNo + vbExclamation, "SOLPED")
                                If r = vbYes Then
                                    Call DataBaseUtils.LogDB("Llegada estimada de SOLPED", "SOLPED: " & spd, "Se registra llegada estimada de SOLPED: " & spd)
                                    campos = Array(CampoDB(TSOLPEDsTrack, spt_usuario), CampoDB(TSOLPEDsTrack, spt_fecha), CampoDB(TSOLPEDsTrack, spt_estado), CampoDB(TSOLPEDsTrack, spt_solped), CampoDB(TSOLPEDsTrack, spt_comentario))
                                    valores = Array(GetCurrentUser, StringUtils.EstablecerFecha("Llegada estimada"), "LLEGADA ESTIMADA", spd, "La llegada estimada es la fecha de registro")
                                    If DataBaseUtils.AddRegister(TSOLPEDsTrack, campos, valores) Then
                                        MsgBox "Registro de fecha estimada exitoso", vbInformation
                                    End If
                                End If
                                MsgBox "La SOLPED: " & spd & " declaró CON OC exitosamente.", vbInformation, "SOLPED Completada"
                            Else
                                MsgBox "Error en registro de SOLPED CON OC. Contacte al Administrador", vbCritical, "Error de Sistema"
                            End If
                        End If
                    End If
                    
                Case "Refacciones"
                    Dim arrResultados() As String
                    Dim i As Long
                    
                    ' Para TODAS las refacciones
                    If Me.AllR.Value = True Then
                        
                        ' 1. Verificamos que haya datos en la lista para evitar Error 9
                        If Me.RefaccionesCods.ListCount = 0 Then
                            Refacciones = Empty
                            Exit Sub
                        End If
                        
                        ' 2. Dimensionamos la memoria para todos los elementos
                        ReDim Refacciones(0 To Me.RefaccionesCods.ListCount - 1)
                        
                        ' 3. Extraemos la columna 0 de toda la lista directamente
                        For i = 0 To Me.RefaccionesCods.ListCount - 1
                            Refacciones(i) = Me.RefaccionesCods.List(i, 0)
                        Next i
                        
                    Else
                        ' Seleccion Multiple o Única
                        If Me.RefaccionesCods.ListCount = 1 Then
                            Refacciones = Array(Me.RefaccionesCods.List(0, 0))
                        Else
                            Refacciones = RefaccionesSeleccionadas()
                        End If
                    End If
                    
                    If IsEmpty(Refacciones) And Me.AllR.Value = False Then
                        MsgBox "Debe seleccionar la(s) refaccione(s)", vbExclamation
                        Exit Sub
                    End If
                    
                    If Me.ConOC.Value Then
                        Call SOLPEDCierre(Refacciones)
                        Unload Me
                        Exit Sub
                    End If
                Case "Equipo"
                    Dim Servicio As Variant
                    Servicio = DataBaseUtils.DatoDataBase(TSOLPEDs, CampoDB(TSOLPEDs, spd_solped), spd, CampoDB(TSOLPEDs, spd_tipo))
                    If IsNull(Servicio) Then
                        Call MsgBox("Falta registro de Tipo en la tabla de SOLPEDs", vbCritical, "StockTech")
                        Exit Sub
                    End If
                    With CostoRefaccion
                        .SOLPEDLabel.Caption = .SOLPEDLabel.Caption & Me.SOLPED.Text
                        If Servicio = "servicio a equipo" Then
                            .TipoLabel.Caption = "Servicio a: " & Me.CodigoEquipo.Caption
                        Else
                            .TipoLabel.Caption = "Equipo: " & Me.CodigoEquipo.Caption
                        End If
                        .Show
                        registroExitoso = .registroSOLPED
                    End With
                    
                    If registroExitoso Then
                        
                        campos = Array(CampoDB(tabla, spt_usuario), CampoDB(tabla, spt_fecha), CampoDB(tabla, spt_estado), CampoDB(tabla, spt_solped), CampoDB(tabla, spt_comentario))
                        ' Usamos la variable obsText que rescatamos en el Paso 1
                        
                        valores = Array(GetCurrentUser(), StringUtils.EstablecerFecha("Generación de OC"), "CON OC", spd, "La generación de la OC es en la fecha de registro")
                        
                        If DataBaseUtils.LogDB("Cierre de SOLPED", "SOLPED: " & spd, "Se registra estado de CON OC de SOLPED: " & spd) Then
                            If DataBaseUtils.AddRegister(tabla, campos, valores) Then
                                'Registro de LLEGADA ESTIMADA
                                r = MsgBox("¿Desea registrar una FECHA ESTIMADA DE LLEGADA?", vbYesNo + vbExclamation, "SOLPED")
                                If r = vbYes Then
                                    Call DataBaseUtils.LogDB("Llegada estimada de SOLPED", "SOLPED: " & spd, "Se registra llegada estimada de SOLPED: " & spd)
                                    campos = Array(CampoDB(TSOLPEDsTrack, spt_usuario), CampoDB(TSOLPEDsTrack, spt_fecha), CampoDB(TSOLPEDsTrack, spt_estado), CampoDB(TSOLPEDsTrack, spt_solped), CampoDB(TSOLPEDsTrack, spt_comentario))
                                    valores = Array(GetCurrentUser, StringUtils.EstablecerFecha("Llegada estimada"), "LLEGADA ESTIMADA", spd, "La llegada estimada es la fecha de registro")
                                    If DataBaseUtils.AddRegister(TSOLPEDsTrack, campos, valores) Then
                                        MsgBox "Registro de fecha estimada exitoso", vbInformation
                                    End If
                                End If
                                
                                MsgBox "La SOLPED: " & spd & " declaró CON OC exitosamente.", vbInformation, "SOLPED Completada"
                                
                            Else
                                MsgBox "Error en registro de SOLPED CON OC. Contacte al Administrador", vbCritical, "Error de Sistema"
                            End If
                        End If
                    End If
                    
            End Select
        End If
    Next
    
    Unload Me
End Sub
Private Sub SOLPEDProceso()
    ' Registro en seguimiento (solo un registro)
    Dim campos As Variant
    Dim valores As Variant
    Dim tabla As String
    
    If Trim(Me.observaciones.Text) = "" Then
        MsgBox "El campo de OBSERVACIONES es OBLIGATORIO", vbExclamation
        Exit Sub
    End If
    
    tabla = TSOLPEDsTrack
    
    campos = Array(CampoDB(tabla, spt_usuario), CampoDB(tabla, spt_fecha), CampoDB(tabla, spt_estado), CampoDB(tabla, spt_solped), CampoDB(tabla, spt_comentario))
    valores = Array(GetCurrentUser, Now(), "PROCESO", Me.SOLPED.Text, Me.observaciones.Text)
    
    If DataBaseUtils.LogDB("SOLPED en Proceso", "SOLPED: " & Me.SOLPED.Text, "Se registra estado de PROCESO de SOLPED: " & Me.SOLPED.Text) Then
        If DataBaseUtils.AddRegister(tabla, campos, valores) Then
            MsgBox "Registro de SOLPED en PROCESO exitoso", vbInformation, "SOLPED en Proceso"
            Unload Me
        Else
            MsgBox "Error en registro de SOLPED. Contacte a Administrador", vbCritical
        End If
    End If
    
    
End Sub
Private Sub SOLPEDLlegada()
    ' Registro en seguimiento (solo un registro)
    Dim campos As Variant
    Dim valores As Variant
    Dim tabla As String
    
    tabla = TSOLPEDsTrack
    Dim revisionRS As ADODB.Recordset
    Dim revision As String
    Dim sql As String
    sql = "SELECT [" & CampoDB(tabla, spt_solped) & "] FROM [" & tabla & "] WHERE [" & CampoDB(tabla, spt_estado) & "] = 'LLEGADA' AND [" & CampoDB(tabla, spt_solped) & "] = '" & Me.SOLPED.Text & "'"
    Set revisionRS = DataBaseUtils.ConsultaSQL(sql, DataBaseUtils.GetDBConnection)
    revision = revisionRS.Fields(CampoDB(tabla, spt_solped)).Value
    On Error Resume Next
    If Not revisionRS Is Nothing Then
        If revisionRS.State <> 0 Then revisionRS.Close
        Set revisionRS = Nothing
    End If
    On Error GoTo 0
    If revision <> Empty Then
        MsgBox "Esta SOLPED ya tiene un registro de LLEGADA", vbExclamation
        Exit Sub
    End If
    
    campos = Array(CampoDB(tabla, spt_usuario), CampoDB(tabla, spt_fecha), CampoDB(tabla, spt_estado), CampoDB(tabla, spt_solped), CampoDB(tabla, spt_comentario))
    valores = Array(GetCurrentUser, EstablecerFecha, "LLEGADA", Me.SOLPED.Text, Me.observaciones.Text)
    
    If DataBaseUtils.LogDB("Llegada de SOLPED", "SOLPED: " & Me.SOLPED.Text, "Se registra fecha de llegada de SOLPED: " & Me.SOLPED.Text) Then
        If DataBaseUtils.AddRegister(tabla, campos, valores) Then
            If DataBaseUtils.SetRegister(TSOLPEDs, CampoDB(TSOLPEDs, spd_llegada), True, CampoDB(TSOLPEDs, spd_solped), Me.SOLPED.Text) Then
                MsgBox "Registro de Llegada de SOLPED exitoso", vbInformation, "Llegada de SOLPED"
            Else
                DataBaseUtils.RollBack tabla
                DataBaseUtils.RollBack TAuditTrail
                
            End If
            Unload Me
        Else
            MsgBox "Error en registro de Llegada SOLPED. Contacte a Administrador", vbCritical
        End If
    End If
    
    
End Sub
Private Sub SOLPEDCierre(ByRef Refacciones As Variant)
    
    Dim spd As String
    Dim ref As String
    Dim i As Long
    Dim r As VbMsgBoxResult
    
    ' Variables para preservar el estado visual
    Dim obsText As String
    Dim solpedCompletamenteCerrada As Boolean
    
    ' =======================================================
    ' PASO 1: CAPTURA DE MEMORIA VIVA
    ' =======================================================
    ' Extraemos la información crítica de la interfaz ANTES de hacer cualquier
    
    spd = Me.SOLPED.Text
    obsText = Me.observaciones.Text
    
    ' Asumimos por defecto que se va a cerrar, a menos que la encontremos viva más adelante
    solpedCompletamenteCerrada = True

    ' =======================================================
    ' PASO 2: MOTOR DE CIERRE TRANSACCIONAL
    ' =======================================================
    For i = LBound(Refacciones) To UBound(Refacciones)
        ref = Refacciones(i)
        
        If Not StringUtils.CerrarRefaccion(ref, spd) Then
            r = MsgBox("Ha ocurrido un error de registro o lo ha cancelado." & Chr(13) & " ¿Desea volver a intentar?", vbYesNo + vbQuestion, "Intentar de Nuevo")
            If r = vbYes Then
                ' Retrocedemos el iterador para reintentar la misma refacción
                i = i - 1
            Else
                Exit Sub
            End If
        End If
    Next
    
    ' =======================================================
    ' PASO 3: RECARGA DE MATRIZ DE PENDIENTES
    ' =======================================================
    ' Disparamos tu función para que vuelva a consultar SQL.
    ' Las SOLPEDs que ya no tengan refacciones pendientes desaparecerán de Me.SOLPED.List.
    Call GetMatrix
    
    ' =======================================================
    ' PASO 4: EVALUACIÓN DE SUPERVIVENCIA (Sustituto Quirúrgico de Match)
    ' =======================================================
    ' En lugar de Application.Match, iteramos la lista reconstruida.
    ' Es matemáticamente exacto y no sufre colisiones de dimensiones (2D vs 1D).
    If Me.SOLPED.ListCount > 0 Then
        For i = 0 To Me.SOLPED.ListCount - 1
            If CStr(Me.SOLPED.List(i, 0)) = spd Then
                ' Si la SOLPED sigue en la lista, significa que le quedan más refacciones.
                ' Cancelamos la bandera de cierre total y abortamos el escaneo.
                solpedCompletamenteCerrada = False
                Exit For
            End If
        Next i
    End If
    
    ' =======================================================
    ' PASO 5: PROTOCOLO DE CIERRE TOTAL
    ' =======================================================
    If solpedCompletamenteCerrada Then
        Dim tabla As String
        Dim campos As Variant
        Dim valores As Variant
        
        tabla = TSOLPEDsTrack
        
        campos = Array(CampoDB(tabla, spt_usuario), CampoDB(tabla, spt_fecha), CampoDB(tabla, spt_estado), CampoDB(tabla, spt_solped), CampoDB(tabla, spt_comentario))
        ' Usamos la variable obsText que rescatamos en el Paso 1
        valores = Array(GetCurrentUser(), EstablecerFecha("Generación de OC"), "CON OC", spd, "La generación de la OC e en la fecha del registro")
        
        If DataBaseUtils.LogDB("Cierre de SOLPED", "SOLPED: " & spd, "Se registra estado de CON OC de SOLPED: " & spd) Then
            If DataBaseUtils.AddRegister(tabla, campos, valores) Then
                r = MsgBox("¿Desea registrar una FECHA ESTIMADA DE LLEGADA?", vbYesNo + vbExclamation, "SOLPED")
                If r = vbYes Then
                    Call DataBaseUtils.LogDB("Llegada estimada de SOLPED", "SOLPED: " & spd, "Se registra llegada estimada de SOLPED: " & spd)
                    campos = Array(CampoDB(TSOLPEDsTrack, spt_usuario), CampoDB(TSOLPEDsTrack, spt_fecha), CampoDB(TSOLPEDsTrack, spt_estado), CampoDB(TSOLPEDsTrack, spt_solped), CampoDB(TSOLPEDsTrack, spt_comentario))
                    valores = Array(GetCurrentUser, StringUtils.EstablecerFecha("Llegada estimada"), "LLEGADA ESTIMADA", spd, "La llegada estimada es la fecha de registro")
                    If DataBaseUtils.AddRegister(TSOLPEDsTrack, campos, valores) Then
                        MsgBox "Registro de fecha estimada exitoso", vbInformation
                    End If
                End If
                MsgBox "La SOLPED: " & spd & " ya no tiene refacciones por registrar y se declarará CERRADA.", vbInformation, "SOLPED Completada"
            Else
                MsgBox "Error en registro de SOLPED CON OC. Contacte al Administrador", vbCritical, "Error de Sistema"
            End If
        End If
    End If
    
    ' =======================================================
    ' PASO 6: DESTRUCCIÓN SEGURA
    ' =======================================================
    ' Una vez concluida toda la lógica y habiendo extraído la información necesaria,
    ' podemos destruir el formulario limpiamente.
    Unload Me

End Sub


Private Sub SOLPED_Change()
    Me.Proceso.Value = False
    Me.ConOC.Value = False
    
    Me.ObservacionLabel.Visible = False
    Me.observaciones.Visible = False
    
    fila = SOLPED.ListIndex
    
    If fila <> -1 Or SOLPED <> Empty Then
        CargarDatosSOLPED (SOLPED.Value)
    End If
End Sub

Private Sub SOLPED_Click()
    Call SOLPED_Change
End Sub

Private Sub UserForm_Activate()
    r = MsgBox("¿Desea registrar llegada de SOLPED's?", vbYesNo)
    If r = vbYes Then
        GetMatrix (True)
    Else
        GetMatrix
    End If
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
    arrSpd = DataBaseUtils.TableDataBase(TSOLPEDs, "*", False, "[" & CampoDB(TSOLPEDs, spd_solped) & "] = '" & idSolpedSeleccionada & "' AND [" & CampoDB(TSOLPEDs, spd_costo) & "] IS NULL")
    
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
