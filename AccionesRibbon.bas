Attribute VB_Name = "AccionesRibbon"
Public edicionTotal As Boolean
Public registroAnt As String
Const MODULE_NAME As String = "AccionesRibbon"
Sub RegistrarSOLPED()
    Const PROC_NAME As String = "RegistrarSOLPED"
    On Error GoTo ErrorRegistro
    Dim Solicitudes As Variant
    Dim numSolicitudes As Long
    Dim shName As String: shName = TSolicitudes
    Dim sh As Worksheet
    If StringUtils.ExisteHoja(shName) Then ' se verifica que exista la hoja
        Set sh = ActiveWorkbook.Worksheets(shName)
        If Seguridad.GetStockTechSheetName(sh) = shName Then ' se verifica que pertenece a la base de datos
            If Not ActiveSheet.Name <> shName Then ' se verifica si la selección ya estaba hecha
                sh.Activate
                If Selection.Rows.Count = 1 Then
                    If Cells(Selection.Row, 7) = False Then
                        IdSolicitud = StringUtils.GetIdFromExcel(TSolicitudes)
                        If IdSolicitud <> 0 Then
                            
                            Dim cod As String
                            Dim tipo As String
                            Dim fila As Integer
                            Dim registrado As Boolean
                            fila = Selection.Row
                            tipo = Cells(fila, 5).Value
                            cod = Cells(fila, 4).Value
                            If cod = "Solicitar Código" Then
                                If tipo = "EQUIPO" Then
                                    With NewEquip
                                        .Show
                                        registrado = .RegistroEquipo
                                    End With
                                    If registrado Then
                                        cod = DataBaseUtils.GetLastIdFromDB(TEquipos, CampoDB(TEquipos, equ_id))
                                        cod = DataBaseUtils.DatoDataBase(TEquipos, CampoDB(TEquipos, equ_id), CInt(cod), CampoDB(TEquipos, equ_codigo))
                                        Cells(fila, 4) = cod
                                        If DataBaseUtils.LogDB("Actualización", "Código de Equipo: " & cod & Chr(13) & " En ID: " & IdSolicitud, "Se realiza registro de código nuevo en la solicitud con ID: " & IdSolicitud) Then
                                            Call DataBaseUtils.UpdateTable(TSolicitudes, True)
                                            Cells(fila, 4).Select '  Se requiere para el funcionamiento inicial de registroSOLPED
                                        Else
                                            Exit Sub
                                        End If
                                        
                                        With registroSOLPED
                                            .CodigoEquipo.Value = cod
                                            .GetMatrix
                                            .RefreshEquipment (True)
                                            .CodigoEquipo.Enabled = False
                                            .Show
                                        End With
                                    End If
                                    Exit Sub
                                End If
                                If tipo = "REFACCIÓN" Then
                                    With NewRefaction
                                        .Show
                                        registrado = .RegistroRefaccion
                                    End With
                                    If registrado Then
                                        cod = DataBaseUtils.GetLastIdFromDB(TRefacciones, CampoDB(TRefacciones, ref_id))
                                        cod = DataBaseUtils.DatoDataBase(TRefacciones, CampoDB(TRefacciones, ref_id), CInt(cod), CampoDB(TRefacciones, ref_material))
                                        Cells(fila, 4) = cod
                                        If DataBaseUtils.LogDB("Actualización", "Código de Refacción: " & cod & Chr(13) & " En ID: " & IdSolicitud, "Se realiza registro de código nuevo en la solicitud con ID: " & IdSolicitud) Then
                                            Call DataBaseUtils.UpdateTable(TSolicitudes, True)
                                            Cells(fila, 4).Select '  Se requiere para el funcionamiento inicial de registroSOLPED
                                        Else
                                            Exit Sub
                                        End If
                                        
                                        With registroSOLPED
                                            .RefaccionCodigo.Value = cod
                                            .GetMatrix
                                            .RefreshRefaction (True)
                                            .Show
                                        End With
                                    End If
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    registroSOLPED.Show
    Exit Sub
ErrorRegistro:
    
    App.SystemError ("Error Registrado: " & Err.Description)
    Exit Sub
End Sub
Sub UpdateSheet()
    Dim shName As String
    shName = Seguridad.GetStockTechSheetName(ActiveSheet)
    If shName = "" Then
        MsgBox "Esta hoja no pertenece a la Base de Datos de StockTech", vbExclamation
        Exit Sub
    End If
    Dim UpdateResult As String
    
    
    
    If edicionTotal Then
        If LogDB("Actualización", "Tabla: " & UCase(shName), "Se realizan cambios en la tabla " & shName, True) Then
            UpdateResult = DataBaseUtils.UpdateTable(shName)
            If UpdateResult = "SIN CAMBIOS" Or UpdateResult = "ERROR" Then
                DataBaseUtils.RollBack TAuditTrail, True
                MsgBox "No se guardará la trazabilidad de cambios nulos en la tabla: " & UCase(shName), vbExclamation, "No AuditTrail"
                Exit Sub
            End If
            If DataBaseUtils.SystemLogDB("Actualización", "Tabla: " & UCase(shName) & " Cambios detectados: " & UpdateResult) Then
                MsgBox "Actualización de la tabla " & UCase(shName) & " exitosa!", vbInformation, "Actualización"
            End If
        End If
        
        Exit Sub
    End If
    
    If LogDB("Actualización", "Tabla: " & UCase(shName), "Se realizan cambios en la tabla " & registroAnt, True) Then
        UpdateResult = DataBaseUtils.UpdateTable(shName)
        If UpdateResult = "SIN CAMBIOS" Or UpdateResult = "ERROR" Then
            DataBaseUtils.RollBack TAuditTrail, True
            MsgBox "No se guardará la trazabilidad de cambios nulos en la tabla: " & UCase(shName), vbExclamation, "No AuditTrail"
            Exit Sub
        End If
        If DataBaseUtils.SystemLogDB("Actualización", "Tabla: " & UCase(shName) & " Registro Anterior: " & registroAnt) Then
            MsgBox "Actualización de la registro en" & UCase(shName) & " exitosa!", vbInformation, "Actualización"
        End If
    End If
    
    App.EnfocarStockTechLabels
End Sub
Sub CalculosPresupuestos()
    ' =======================================================
    ' 1. DECLARACIÓN DE VARIABLES Y ENTORNO
    ' =======================================================
    Dim tabla As Variant, campos As Variant
    Dim tablaName As String, nombreTablaOficial As String
    Dim columnaNva1 As String, columnaNva2 As String
    Dim i As Long
    
    ' Variables para el gráfico
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim chtObj As ChartObject
    Dim cht As Chart
    Dim Serie As Series
    Dim nombreGrafico As String
    
    tablaName = "Prestos. Calculados"
    ' Nombre que Excel le dará a la tabla (reemplaza espacios por guiones bajos)
    nombreTablaOficial = Replace(tablaName, " ", "_")
    
    columnaNva1 = "Presupuesto Restante"
    columnaNva2 = "Porcentaje Utilizado"
    nombreGrafico = "Porcentaje Utilizado"
    
    ' =======================================================
    ' 2. OBTENCIÓN Y RENDERIZADO DE DATOS
    ' =======================================================
    tabla = DataBaseUtils.TableDataBase(TPresupuestos, "*")
    campos = TablasDB.CamposTablaDB(TPresupuestos) ' Corregido para usar tu fábrica
    
    ' Renderizamos la tabla en Excel
    DataBaseUtils.GetExcelTable tablaName, tabla, campos, refresh:=True
    
    ' Una vez renderizada, asignamos la hoja activa a nuestra variable de entorno
    Set ws = ActiveSheet
    
    ' =======================================================
    ' 3. INYECCIÓN DE COLUMNAS DE CÁLCULO
    ' =======================================================
    ' Siempre es más seguro referenciar ws.Cells que Cells crudo
    ws.Cells(1, 6).Value = columnaNva1
    ws.Cells(1, 6).WrapText = True
    ws.Cells(1, 7).Value = columnaNva2
    
    ' Iteración de cálculos. Se usa <> "" para evaluar celdas vacías con precisión.
    i = 2
    While ws.Cells(i, 3).Value <> ""
        ' Cálculo de Saldo con casteo de Moneda
        ws.Cells(i, 6).Value = CCur(ConsultasSQL.GetSaldoCuenta(CStr(ws.Cells(i, 3).Value)))
        
        ' Cálculo de Porcentaje (Se añade IFERROR para evitar colapso si Presupuesto es 0)
        ws.Cells(i, 7).FormulaR1C1 = "=IFERROR(100*((RC[-2]-RC[-1])/RC[-2]), 0)"
        
        i = i + 1
    Wend
    
    ws.Columns.AutoFit
    
    ' =======================================================
    ' 4. PREPARACIÓN DEL OBJETO LISTOBJECT (La Tabla Oficial)
    ' =======================================================
    On Error Resume Next
    Set tbl = ws.ListObjects(nombreTablaOficial)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "Error de Arquitectura: No se encontró la tabla interna: " & nombreTablaOficial, vbCritical
        Exit Sub
    End If
    
    ' =======================================================
    ' 5. CONSTRUCCIÓN DEL GRÁFICO (El Contenedor)
    ' =======================================================
    ' Purgamos gráficos anteriores
    On Error Resume Next
    ws.ChartObjects(nombreGrafico).Delete
    On Error GoTo 0
    
    Set chtObj = ws.ChartObjects.Add( _
        Left:=tbl.Range.Left + tbl.Range.Width + 20, _
        Top:=tbl.Range.Top, _
        Width:=450, _
        Height:=250)
        
    chtObj.Name = nombreGrafico
    Set cht = chtObj.Chart
    cht.ChartType = xlColumnClustered
    
    ' =======================================================
    ' 6. ENSAMBLAJE DE SERIES DE DATOS
    ' =======================================================
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    
    Set Serie = cht.SeriesCollection.NewSeries
    With Serie
        .Values = tbl.ListColumns(columnaNva2).DataBodyRange
        .XValues = tbl.ListColumns(CampoDB(TPresupuestos, pre_id_cuenta)).DataBodyRange
        .Name = columnaNva2
    End With
    
    ' =======================================================
    ' 7. DISEÑO, UX Y BLINDAJE DE COLUMNAS OCULTAS
    ' =======================================================
    With cht
        .HasTitle = True
        .ChartTitle.Text = columnaNva2
        .HasLegend = False
        .Axes(xlCategory).TickLabels.Orientation = 45
    End With
    
    ' ¡CRÍTICO!: Obligamos a la gráfica a renderizar datos aunque ocultes sus columnas después
    chtObj.Chart.PlotVisibleOnly = False
    
    ' =======================================================
    ' 8. CIERRE Y SEGURIDAD (Limpieza de Interfaz)
    ' =======================================================
    ' Ocultamos las columnas que no necesita ver el usuario
    campos = Array(CampoDB(TPresupuestos, pre_id_cuenta), CampoDB(TPresupuestos, pre_presupuesto), columnaNva1)
    Call StringUtils.SeleccionColumnasTabla(tablaName, campos)
    
    Seguridad.UnmarkStockTechSheet ws
    'Seguridad.LockSheet (ActiveSheet)
    App.EnfocarStockTechLabels
    
End Sub
Sub MostrarSolicitudes()
    Call DataBaseUtils.GetExcelTable(TSolicitudes, refresh:=True)
    campos = CamposTablaDB(TSolicitudes)
    Call StringUtils.DropFromArray(campos, 0)
    Call StringUtils.SeleccionColumnasTabla(TSolicitudes, campos)
    Call StringUtils.FiltrarTabla(TSolicitudes, CampoDB(TSolicitudes, sol_realizada), False)
    Call Seguridad.LockSheet(ActiveSheet)
    App.EnfocarStockTechLabels
End Sub
Sub EliminarRegistro()
    Dim tabla As String
    tabla = Seguridad.GetStockTechSheetName(ActiveSheet)
    If tabla = Empty Then
        MsgBox "La Hoja seleccionada no pertenece a StoskTech"
        Exit Sub
    End If
    If tabla = TAuditTrail Then
        MsgBox "Los registros de TRAZABILIDAD no pueden ser eliminados", vbCritical, "Eliminación no Válida"
        Exit Sub
    End If
    Dim id As Integer
    Dim fila As Long
    If Selection.Rows.Count > 1 Then
        MsgBox "Por seguridad, Solo se puede eliminar un registro a la vez", vbExclamation, "Eliminación no Válida"
        Exit Sub
    End If
    fila = Selection.Row
    id = CInt(Cells(fila, 1).Value)
    Dim campos As Variant
    campos = CamposTablaDB(tabla)
    Dim registro As String
    registro = ""
    For i = 0 To UBound(campos)
        registro = registro & campos(i) & ": " & Cells(fila, i + 1) & "  " & Chr(13)
    Next
    Call Left(registro, Len(registro) - 4)
    
    If id <> 0 Then
        If DataBaseUtils.LogDB("Elimnación", "Registro con ID: " & id, "Se elimna Registro: " & registro) Then
            If DataBaseUtils.DeleteRegisterByID(tabla, id) Then
                MsgBox "Eliminación exitosa", vbInformation, "Eliminación en DB"
                DataBaseUtils.GetExcelTable tabla, refresh:=True
            Else
                RollBack (TAuditTrail)
            End If
        End If
        Exit Sub
    End If
    MsgBox "Seleccione un registro válido", vbInformation, "Eliminación No Válida"
End Sub
Sub EditarRegistro()
    edicionTotal = False
    Dim tabla As String
    tabla = Seguridad.GetStockTechSheetName(ActiveSheet)
    If tabla = Empty Then
        MsgBox "La Hoja seleccionada no pertenece a StoskTech"
        Exit Sub
    End If
    If tabla = TAuditTrail Then
        MsgBox "Los registros de TRAZABILIDAD no pueden ser modificados", vbCritical, "Modificación no Válida"
        Exit Sub
    End If
    Dim id As Integer
    Dim fila As Long
    If Selection.Rows.Count > 1 Then
        If edicionTotal = False Then
            MsgBox "Por seguridad, Solo se puede editar un registro a la vez", vbExclamation, "Modificación no Válida"
        End If
        Exit Sub
    End If
    edicionTotal = Seguridad.VerificarPrivilegio(GetCurrentUser, "Editar Todo")
    If edicionTotal Then
        r = MsgBox("¿Desea activar la edición de la tabla completa?", vbYesNo + vbQuestion, "Edición de tabla")
        If r = vbYes Then
            ActiveSheet.Cells.Locked = False
            MsgBox "Puede editar la Hoja", vbInformation, "Edición masiva"
            Exit Sub
        End If
    End If
    fila = Selection.Row
    id = CInt(Cells(fila, 1).Value)
    Dim campos As Variant
    campos = CamposTablaDB(tabla)
    
    If id <> 0 And Selection.Rows.Count = 1 Then
        registroAnt = ""
        For i = 0 To UBound(campos)
            registroAnt = registroAnt & campos(i) & ": " & Cells(fila, i + 1) & "  " & Chr(13)
        Next
        Call Left(registroAnt, Len(registroAnt) - 8)
        
        Dim RG As Range
        
        Set RG = ActiveSheet.Range(ActiveSheet.Cells(fila, 2), ActiveSheet.Cells(fila, UBound(campos) + 1))
        
        With RG
            .Locked = False
            ' color amarillo
            .Interior.Color = RGB(255, 255, 0)
            .Font.Color = RGB(0, 0, 0)
        End With
        MsgBox "Registro habilitado para edición. Terminando actualice la tabla"
        Exit Sub
    End If
    MsgBox "Seleccione un registro válido", vbInformation, "Modificación No Válida" ' Ruta para la edición de un ID inválido
End Sub
Public Sub VerUserDetails(ByVal user As String)
    Dim tabla As Variant
    Dim ws As Worksheet
    Dim tablaName As String
    
    ' =======================================================
    ' 1. OBTENCIÓN DE DATOS BASE (spd_cuenta ya incluido en el JOIN)
    ' =======================================================
    tabla = ConsultasSQL.GetTablaDetalleUsuario(user)
    
    If IsEmpty(tabla) Then
        MsgBox "El usuario " & UCase(user) & " no tiene datos disponibles.", vbExclamation, "Detalles del Usuario"
        Exit Sub
    End If
    
    tablaName = user & " Details"
    
    ' =======================================================
    ' 2. RENDERIZADO VISUAL
    ' =======================================================
    Call DataBaseUtils.GetExcelTable(tablaName, tabla(0), tabla(1), True)
    Set ws = ActiveSheet
    
    ' =======================================================
    ' 3. COLUMNAS VISIBLES EN LA TABLA DE DETALLE
    '    Mostramos todo excepto el ID (índice 0)
    ' =======================================================
    Dim camposDetalle As Variant
    camposDetalle = tabla(1)
    Call StringUtils.DropFromArray(camposDetalle, 0) ' Quitamos el ID
    Call StringUtils.SeleccionColumnasTabla(tablaName, camposDetalle)
    
    ' =======================================================
    ' 4. RESUMEN POR CUENTA (separado visualmente)
    '    colInicio apunta a una zona segura DESPUÉS de la tabla de detalle
    ' =======================================================
    Dim colResumen As Long
    colResumen = UBound(tabla(1)) + 3  ' +1 por base 0, +2 de margen visual
    
    Call ConsultasSQL.CostByCountFromUser(ws, user, colResumen)
    
    ' =======================================================
    ' 5. CIERRE, SEGURIDAD Y UX
    ' =======================================================
    Call Seguridad.UnmarkStockTechSheet(ws)
    Call App.EnfocarStockTechLabels
End Sub
Sub AprobarSOLPED()
    r = MsgBox("¿Desea ver la tabla de SOLPED's pendientes por aprobación?", vbYesNo, "StockTech Aprobación")
    If r = vbYes Then
        DataBaseUtils.GetExcelTable TSOLPEDs, refresh:=True
        Call StringUtils.FiltrarTabla(TSOLPEDs, CampoDB(TSOLPEDs, spd_aprobacion), False)
        Dim camposNoVis As Variant
        camposNoVis = Array(spd_id_solicitud, spd_comentario, spd_costo, spd_no_factura, spd_orden_compra)
        Call StringUtils.OcultarColumnasTabla(TSOLPEDs, camposNoVis)
        Exit Sub
    End If
    AprobacionSOLPED.Show
End Sub
