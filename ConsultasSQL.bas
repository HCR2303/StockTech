Attribute VB_Name = "ConsultasSQL"
Option Private Module

Const MODULE_NAME As String = "ConsultasSQL"
Public Function GetSaldoCuenta(ByVal NumeroCuenta As String) As Double
    Const PROC_NAME As String = "GetSaldoCuenta"
    Dim rs As Object
    Dim SQL As String
    Dim presupuestoBase As Double
    Dim totalGastado As Double
    Dim camposPresupuestos As Variant
    camposPresupuestos = CamposTablaDB(TPresupuestos)
    
    Dim camposSOLPEDs As Variant
    camposSOLPEDs = CamposTablaDB(TSOLPEDs)
    
    SQL = "SELECT [" & camposPresupuestos(pre_presupuesto) & "], " & _
          "(SELECT SUM([" & camposSOLPEDs(spd_costo) & "]) FROM [" & TSOLPEDs & "] " & _
          "WHERE [" & camposSOLPEDs(spd_cuenta) & "] = [" & TPresupuestos & "].[" & camposPresupuestos(pre_id_cuenta) & "]) AS [TotalCosto] " & _
          "FROM [" & TPresupuestos & "] WHERE [" & camposPresupuestos(pre_id_cuenta) & "] = '" & NumeroCuenta & "'"
    
    On Error GoTo ErrorHandler
    Set rs = DataBaseUtils.ConsultaSQL(SQL, DataBaseUtils.GetDBConnection)
    
    If Not rs.EOF Then
        presupuestoBase = IIf(IsNull(rs.Fields(camposPresupuestos(pre_presupuesto)).Value), 0, rs.Fields(camposPresupuestos(pre_presupuesto)).Value)
        totalGastado = IIf(IsNull(rs.Fields("TotalCosto").Value), 0, rs.Fields("TotalCosto").Value)
        
        GetSaldoCuenta = presupuestoBase - totalGastado
    Else
        MsgBox "La consulta se ejecutó, pero no se encontró la cuenta: " & NumeroCuenta, vbExclamation
        GetSaldoCuenta = 0
    End If
    
limpieza:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Function

ErrorHandler:
    App.SystemError "Error en [" & PROC_NAME & "] al calcular saldo de: " & NumeroCuenta & " - " & Err.Description
    GetSaldoCuenta = 0
    Resume limpieza
End Function

Public Function GetTablaDetalleUsuario(ByVal nombreUsuario As String) As Variant
    Const PROC_NAME As String = "GetTablaDetalleUsuario"
    Dim rs As Object
    Dim SQL As String
    
    Dim arrSol As Variant
    Dim arrSpd As Variant
    Dim arrHeaders As Variant
    
    arrSol = TablasDB.CamposTablaDB(TablasDB.TSolicitudes)
    arrSpd = TablasDB.CamposTablaDB(TablasDB.TSOLPEDs)
    
    ' *** spd_cuenta añadido en posición 6, Estado_Operativo pasa a 7 ***
    arrHeaders = Array( _
        arrSol(CNameSolicitudes.sol_id_solicitud), _
        arrSol(CNameSolicitudes.sol_fecha), _
        arrSol(CNameSolicitudes.sol_codigo), _
        arrSpd(CNameSOLPEDs.SPD_SOLPED), _
        arrSpd(CNameSOLPEDs.SPD_REFACCION), _
        arrSpd(CNameSOLPEDs.spd_costo), _
        arrSpd(CNameSOLPEDs.spd_cuenta), _
        "Estado_Operativo" _
    )
    
    SQL = "SELECT Sol.[" & arrHeaders(0) & "], " & _
                 "Sol.[" & arrHeaders(1) & "], " & _
                 "Sol.[" & arrHeaders(2) & "], " & _
                 "Spd.[" & arrHeaders(3) & "], " & _
                 "Spd.[" & arrHeaders(4) & "], " & _
                 "Spd.[" & arrHeaders(5) & "], " & _
                 "Spd.[" & arrHeaders(6) & "], " & _
                 "IIf(IsNull(Spd.[" & arrHeaders(5) & "]) OR Spd.[" & arrHeaders(5) & "] = 0, 'Pendiente', 'Cerrado') AS [" & arrHeaders(7) & "] " & _
          "FROM ([" & TablasDB.TSolicitudes & "] AS Sol " & _
          "INNER JOIN [" & TablasDB.TSOLPEDs & "] AS Spd ON Sol.[" & arrHeaders(0) & "] = Spd.[" & arrSpd(CNameSOLPEDs.spd_id_solicitud) & "]) " & _
          "WHERE Sol.[" & arrSol(CNameSolicitudes.sol_usuario) & "] = '" & Replace(nombreUsuario, "'", "''") & "' " & _
            "AND Sol.[" & arrSol(CNameSolicitudes.sol_realizada) & "] = True " & _
          "ORDER BY Sol.[" & arrHeaders(1) & "] DESC"

    On Error GoTo ErrorHandler
    Set rs = DataBaseUtils.ConsultaSQL(SQL, DataBaseUtils.GetDBConnection)
    
    If Not rs.EOF Then
        GetTablaDetalleUsuario = Array(rs.GetRows, arrHeaders)
    Else
        GetTablaDetalleUsuario = Empty
    End If
    
limpieza:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Function

ErrorHandler:
    App.SystemError "Error en [" & PROC_NAME & "] al generar tabla para: " & nombreUsuario & " - " & Err.Description
    GetTablaDetalleUsuario = Empty
    Resume limpieza
End Function
Public Function FormatoSQL(ByVal dato As Variant) As String
    ' Esta función recibe cualquier cosa y devuelve el texto exacto para inyectar en Access
    Select Case VarType(dato)
        Case vbBoolean
            FormatoSQL = IIf(dato, "-1", "0")
        Case vbDate
            FormatoSQL = "#" & Format(dato, "yyyy-mm-dd HH:mm:ss") & "#"
        Case vbString
            If Len(Trim(dato)) = 0 Then
                FormatoSQL = "NULL"
            Else
                FormatoSQL = "'" & Replace(dato, "'", "''") & "'"
            End If
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            If IsNumeric(dato) Then
                FormatoSQL = Replace(CStr(dato), ",", ".")
            Else
                FormatoSQL = "NULL"
            End If
        Case Else
            FormatoSQL = "NULL"
    End Select
End Function
Public Sub CostByCountFromUser(ByVal ws As Worksheet, ByVal user As String, ByVal colInicio As Long)
    
    Dim SQL_Resumen As String
    Dim matrizResumen As Variant
    Dim f As Long
    Dim chtObj As ChartObject
    Dim cht As Chart
    Dim Serie As Series
    Dim nombreGrafico As String
    Dim ultimaFilaResumen As Long
    
    ' Variables para las fábricas de nombres
    Dim arrSol As Variant
    Dim arrSpd As Variant
    
    nombreGrafico = "GraficoCostos_" & user
    
    ' =======================================================
    ' PASO 1: ENSAMBLAJE DE LA CONSULTA SQL (JOIN + GROUP BY)
    ' =======================================================
    ' Instanciamos los vectores de columnas usando tu estándar
    arrSol = TablasDB.CamposTablaDB(TablasDB.TSolicitudes)
    arrSpd = TablasDB.CamposTablaDB(TablasDB.TSOLPEDs)
    
    ' Construimos la petición replicando la relación de tu tabla de detalles
    SQL_Resumen = "SELECT Spd.[" & arrSpd(CNameSOLPEDs.spd_cuenta) & "], SUM(Spd.[" & arrSpd(CNameSOLPEDs.spd_costo) & "]) " & _
                  "FROM ([" & TablasDB.TSolicitudes & "] AS Sol " & _
                  "INNER JOIN [" & TablasDB.TSOLPEDs & "] AS Spd ON Sol.[" & arrSol(CNameSolicitudes.sol_id_solicitud) & "] = Spd.[" & arrSpd(CNameSOLPEDs.spd_id_solicitud) & "]) " & _
                  "WHERE Sol.[" & arrSol(CNameSolicitudes.sol_usuario) & "] = '" & Replace(user, "'", "''") & "' " & _
                    "AND Sol.[" & arrSol(CNameSolicitudes.sol_realizada) & "] = True " & _
                  "GROUP BY Spd.[" & arrSpd(CNameSOLPEDs.spd_cuenta) & "]"
                  
    ' =======================================================
    ' PASO 2: EXTRACCIÓN Y RENDERIZADO OCULTO
    ' =======================================================
    matrizResumen = DataBaseUtils.GetFromSQL(SQL_Resumen) ' Usamos tu motor de extracción
    
    If Not IsArray(matrizResumen) Then
        ' Si no hay registros válidos para sumar, abortamos la creación del gráfico
        Exit Sub
    End If
    
    ' Imprimimos los encabezados en la zona de operaciones (Ej. Columna K)
    ws.Cells(1, colInicio).Value = "Cuenta"
    ws.Cells(1, colInicio + 1).Value = "Costo Total"
    
    ' Volcamos la matriz procesada
    For f = 0 To UBound(matrizResumen, 2)
        ' Si la cuenta llega nula, le ponemos un texto genérico para que el gráfico no falle
        If IsNull(matrizResumen(0, f)) Then
            ws.Cells(f + 2, colInicio).Value = "Sin Asignar"
        Else
            ws.Cells(f + 2, colInicio).Value = CStr(matrizResumen(0, f))
        End If
        
        ' Forzamos el casteo a Moneda, asegurando que los nulos se vuelvan 0
        If IsNull(matrizResumen(1, f)) Then
            ws.Cells(f + 2, colInicio + 1).Value = CCur(0)
        Else
            ws.Cells(f + 2, colInicio + 1).Value = CCur(matrizResumen(1, f))
        End If
    Next f
    
    ultimaFilaResumen = UBound(matrizResumen, 2) + 2
    
    ' =======================================================
    ' PASO 3: PURGADO DE GRÁFICOS PREVIOS
    ' =======================================================
    On Error Resume Next
    ws.ChartObjects(nombreGrafico).Delete
    On Error GoTo 0
    
    ' =======================================================
    ' PASO 4: CONSTRUCCIÓN DEL CONTENEDOR VISUAL
    ' =======================================================
   ' Reemplazar solo el bloque PASO 4 (construcción del contenedor visual):

    Set chtObj = ws.ChartObjects.Add( _
        Left:=ws.Cells(1, colInicio).Left, _
        Top:=ws.Cells(UBound(matrizResumen, 2) + 4, colInicio).Top, _
        Width:=350, _
        Height:=220)
        
    chtObj.Name = nombreGrafico
    
    Set cht = chtObj.Chart
    cht.ChartType = xlColumnClustered
    
    ' =======================================================
    ' PASO 5: ENSAMBLAJE QUIRÚRGICO DE SERIES
    ' =======================================================
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    
    Set Serie = cht.SeriesCollection.NewSeries
    With Serie
        .XValues = ws.Range(ws.Cells(2, colInicio), ws.Cells(ultimaFilaResumen, colInicio))
        .Values = ws.Range(ws.Cells(2, colInicio + 1), ws.Cells(ultimaFilaResumen, colInicio + 1))
        .Name = "Costo por Cuenta"
    End With
    
    ' =======================================================
    ' PASO 6: UX Y BLINDAJE
    ' =======================================================
    With cht
        .HasTitle = True
        .ChartTitle.Text = "Distribución de Costos"
        .HasLegend = False
        .Axes(xlCategory).TickLabels.Orientation = 45
        .Axes(xlValue).TickLabels.NumberFormat = "$#,##0.00"
    End With
    
    cht.PlotVisibleOnly = False
    
    ws.Columns(colInicio).Hidden = True
    ws.Columns(colInicio + 1).Hidden = True

End Sub

