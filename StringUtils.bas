Attribute VB_Name = "StringUtils"
Option Private Module
Const MODULE_NAME As String = "StringUtils"
Public Function CrearCamposSQLTablas(ByVal CamposP As Variant, Optional ByVal camposU As Variant) As String
    Dim cU As String, cP As String
    
    ' 1. Manejar Tabla U (La que manda - Parámetro Opcional)
    If IsMissing(camposU) Or IsEmpty(camposU) Then
        cU = "U.*"
    ElseIf Not IsArray(camposU) Then
        ' Si es texto, validamos que no sea un asterisco ni un string vacío oculto
        If Trim(CStr(camposU)) = "*" Or Trim(CStr(camposU)) = "" Then
            cU = "U.*"
        Else
            cU = "U.[" & camposU & "]"
        End If
    Else
        ' El uso de Join aquí es brillante, lo mantenemos intacto
        cU = "U.[" & Join(camposU, "], U.[") & "]"
    End If
    
    ' 2. Manejar Tabla P (Campos específicos - Parámetro Obligatorio)
    If IsEmpty(CamposP) Then ' Cambiamos IsMissing por IsEmpty porque no es Optional
        cP = "P.*"
    ElseIf Not IsArray(CamposP) Then
        If Trim(CStr(CamposP)) = "*" Or Trim(CStr(CamposP)) = "" Then
            cP = "P.*"
        Else
            cP = "P.[" & CamposP & "]"
        End If
    Else
        cP = "P.[" & Join(CamposP, "], P.[") & "]"
    End If
    
    CrearCamposSQLTablas = cU & ", " & cP
End Function
Function CrearCamposSQL(campos As Variant) As String
    Dim cU As String, cP As String
    
    If Not IsArray(campos) Then
        If campos = "*" Then cU = "*" Else cU = "[" & campos & "]"
    Else
        cU = "[" & Join(campos, "], [") & "]"
    End If
    
    CrearCamposSQL = cU
End Function
Public Sub UltimoDatoTabla(ByVal tabla As String, ByVal columnaID As Variant)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim arrCampos As Variant
    Dim nombreColumna As String
    Dim colIndex As Long
    Dim maxID As Double
    
    On Error GoTo ErrorHandler
    
    ' 1. Conexión segura con la hoja de Excel
    Set ws = ActiveWorkbook.Worksheets(tabla)
    If ws.ListObjects.Count = 0 Then Exit Sub
    Set lo = ws.ListObjects(1)
    
    ' 2. LIMPIEZA ESTRICTA DE FILTROS DE LA TABLA (El paso crucial)
    ' Intervenimos directamente el AutoFilter del ListObject, no el de la Hoja
    If Not lo.AutoFilter Is Nothing Then
        If lo.AutoFilter.FilterMode Then
            lo.AutoFilter.ShowAllData
        End If
    End If
    
    ' 3. Traducción Automática de Enum a Nombre Real
    If IsNumeric(columnaID) Then
        arrCampos = TablasDB.CamposTablaDB(tabla)
        If columnaID >= LBound(arrCampos) And columnaID <= UBound(arrCampos) Then
            nombreColumna = arrCampos(columnaID)
        Else
            Call App.SystemError("Fallo de Arquitectura: El Enum del ID supera las columnas de [" & tabla & "].")
            Exit Sub
        End If
    Else
        nombreColumna = CStr(columnaID)
    End If
    
    ' 4. Validación de Tabla Vacía
    If lo.DataBodyRange Is Nothing Then Exit Sub
    
    ' 5. Búsqueda exacta de la columna
    On Error Resume Next
    colIndex = lo.ListColumns(nombreColumna).Index
    On Error GoTo ErrorHandler
    
    If colIndex = 0 Then
        Call App.SystemError("No se encontró la cabecera '" & nombreColumna & "' en la tabla de Excel.")
        Exit Sub
    End If
    
    ' 6. Extracción del valor máximo (Ahora sí, con toda la data expuesta)
    maxID = Application.WorksheetFunction.Max(lo.ListColumns(colIndex).DataBodyRange)
    
    ' 7. Aplicación del Filtro
    Application.ScreenUpdating = False
    
    lo.Range.AutoFilter Field:=colIndex, Criteria1:=maxID
    
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Call App.SystemError("Error en UltimoDatoTabla: " & Err.Description)
End Sub
Public Sub DropFromArray(ByRef matriz As Variant, ByVal indexToRemove As Long)
    Dim i As Long
    Dim lowerBound As Long
    Dim upperBound As Long
    
    
    If Not IsArray(matriz) Then Exit Sub
    
    lowerBound = LBound(matriz)
    upperBound = UBound(matriz)
    
    
    If indexToRemove < lowerBound Or indexToRemove > upperBound Then Exit Sub
    
    If lowerBound = upperBound Then
        matriz = Empty
        Exit Sub
    End If
    
    For i = indexToRemove To upperBound - 1
        matriz(i) = matriz(i + 1)
    Next i
    
    ReDim Preserve matriz(lowerBound To upperBound - 1)
End Sub
Public Sub SeleccionColumnasTabla(ByVal tabla As String, ByVal camposVisibles As Variant, Optional UltimoValor As Boolean = False)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim arrCampos As Variant
    Dim nombreColumna As String
    
    On Error GoTo ErrorHandler
    
    
    Application.ScreenUpdating = False
    
    Set ws = ActiveWorkbook.Worksheets(tabla)
    
    
    If ws.ListObjects.Count = 0 Then GoTo SalidaSegura
    Set lo = ws.ListObjects(1)
    
    If lo.ShowAutoFilter Then
        If ws.FilterMode Then ws.ShowAllData
    End If
    
    ' 3. Ocultamos toda la tabla de un solo golpe
    lo.Range.EntireColumn.Hidden = True
    
    ' 4. Cargamos el diccionario de nombres de esta tabla específica
    arrCampos = TablasDB.CamposTablaDB(tabla)
    
    ' 5. Iteración y Traducción
    For i = LBound(camposVisibles) To UBound(camposVisibles)
        
        ' Evaluamos si el dato es un número (Enum) o un Texto (Nombre directo)
        If IsNumeric(camposVisibles(i)) Then
            ' Escudo de límites: Verificamos que el Enum exista en la tabla
            If camposVisibles(i) >= LBound(arrCampos) And camposVisibles(i) <= UBound(arrCampos) Then
                nombreColumna = arrCampos(camposVisibles(i))
            Else
                nombreColumna = ""
            End If
        Else
            nombreColumna = CStr(camposVisibles(i))
        End If
        
        ' 6. Mostrar la columna por su NOMBRE EXACTO (A prueba de movimientos físicos en Excel)
        If nombreColumna <> "" Then
            ' Encendemos On Error Resume Next SOLO para esta línea, por si
            ' el nombre en el diccionario no coincide con el de la hoja de Excel
            On Error Resume Next
            lo.ListColumns(nombreColumna).Range.EntireColumn.Hidden = False
            On Error GoTo ErrorHandler ' Restauramos el manejo de errores global
        End If
        
    Next i
    
    
    If UltimoValor Then Call UltimoDatoTabla(tabla, 0)
    ws.Select
    
    Application.Goto Reference:=lo.HeaderRowRange(1, 1), Scroll:=True
SalidaSegura:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Call App.SystemError("Error en SeleccionColumnasTabla: " & Err.Description)
    Resume SalidaSegura
End Sub
Public Sub OcultarColumnasTabla(ByVal tabla As String, ByVal CamposAOcultar As Variant)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim arrCampos As Variant
    Dim nombreColumna As String
    
    ' =======================================================
    ' PASO 1: ESCUDO DE INTERRUPCIÓN Y CONGELAMIENTO VISUAL
    ' =======================================================
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    Set ws = ActiveWorkbook.Worksheets(tabla)
    
    ' Validamos que la hoja tenga al menos una tabla oficial para evitar Error 9
    If ws.ListObjects.Count = 0 Then GoTo SalidaSegura
    Set lo = ws.ListObjects(1)
    
    ' =======================================================
    ' PASO 2: CARGA DEL DICCIONARIO DE DATOS
    ' =======================================================
    ' Extraemos la matriz de encabezados originales de la base de datos
    arrCampos = TablasDB.CamposTablaDB(tabla)
    
    ' =======================================================
    ' PASO 3: ITERACIÓN Y TRADUCCIÓN HÍBRIDA
    ' =======================================================
    For i = LBound(CamposAOcultar) To UBound(CamposAOcultar)
        
        ' Evaluamos si el desarrollador inyectó un número (Enum) o un Texto
        If IsNumeric(CamposAOcultar(i)) Then
            
            ' Escudo de límites: Aseguramos que el número exista en el diccionario
            If CamposAOcultar(i) >= LBound(arrCampos) And CamposAOcultar(i) <= UBound(arrCampos) Then
                nombreColumna = arrCampos(CamposAOcultar(i))
            Else
                nombreColumna = ""
            End If
            
        Else
            ' Si es texto, lo tomamos de manera literal
            nombreColumna = CStr(CamposAOcultar(i))
        End If
        
        ' =======================================================
        ' PASO 4: OCULTAMIENTO QUIRÚRGICO FÍSICO
        ' =======================================================
        If nombreColumna <> "" Then
            ' Encendemos On Error Resume Next estrictamente para esta línea.
            ' Si la columna ya fue borrada o no existe, la macro no colapsará.
            On Error Resume Next
            lo.ListColumns(nombreColumna).Range.EntireColumn.Hidden = True
            On Error GoTo ErrorHandler ' Restauramos el escudo global inmediatamente
        End If
        
    Next i

SalidaSegura:
    ' =======================================================
    ' PASO 5: RESTAURACIÓN DE INTERFAZ Y SALIDA
    ' =======================================================
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    ' Trazabilidad del error hacia el motor principal
    Call App.SystemError("Error en OcultarColumnasTabla: " & Err.Description)
    Resume SalidaSegura
End Sub
Public Function ExisteTabla(ByVal nombreTabla As String) As Boolean
    Dim lo As ListObject
    On Error Resume Next
    
    Set lo = ActiveWorkbook.Range(nombreTabla).ListObject
    
    On Error GoTo 0
    ExisteTabla = Not lo Is Nothing
End Function
Public Function ExisteHoja(ByVal nombreHoja As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    
    Set ws = ActiveWorkbook.Worksheets(nombreHoja)
    
    On Error GoTo 0
    ExisteHoja = Not ws Is Nothing
End Function
Public Function GetIdFromExcel(tabla As String) As Integer
    
    Dim idColumnDBName As String
    Dim fila As Integer
    Dim idBuscado As Integer: idBuscado = 0
    
    ' Si la selección es valida se continua
    If Selection.Rows.Count = 1 Then
        On Error Resume Next
        fila = Cells(ActiveCell.Row, 1)
        idColumnDBName = DataBaseUtils.GetColumnName(tabla, 0)
        idBuscado = DataBaseUtils.DatoDataBase(tabla, idColumnDBName, fila, idColumnDBName)
        GetIdFromExcel = idBuscado
        On Error GoTo 0
    Else
        App.SystemError ("No hay selección valida desde GetIdFromExcel")
        GetIdFromExcel = idBuscado
    End If
    
End Function
Public Function EstablecerFecha(Optional ByVal tipoFecha As String = "") As Date
    
    If tipoFecha <> "" Then
        MsgBox "Establezca la fecha para " & UCase(tipoFecha)
        Calendario.Caption = UCase(tipoFecha)
    End If
    
    Calendario.Show
    While Calendario.fechaSeleccionada = Empty
        MsgBox "La selección de una fecha es obligatorio", vbExclamation, "Registro de fecha"
        On Error Resume Next
            Calendario.Show
        On Error GoTo 0
    Wend
    EstablecerFecha = Calendario.fechaSeleccionada
    Unload Calendario
End Function
Public Function CerrarRefaccion(ByVal Refaccion As String, ByVal SOLPED As String) As Boolean
    Dim exito As Boolean
    
    With CostoRefaccion
        .SOLPEDLabel = .SOLPEDLabel.Caption & SOLPED
        .TipoLabel = "Refacción: " & Refaccion
        .Show
        exito = .registroSOLPED
    End With
    If exito Then
        'Registro de AuditTrail
        If Not DataBaseUtils.LogDB("Cambio de estado a CERRADA", "Refacción: " & Refaccion, _
                        "Refacción: " & Refaccion & " con referencia a la SOLPED: " & SOLPED & " CERRADA") Then
            MsgBox "Error de registro de trazabilidad. Registro comprometido", vbCritical, "Error en AuditTrail"
            Exit Function
        End If
        CerrarRefaccion = True
    End If
End Function

Public Function CampoDB(ByVal nombreTabla As String, ByVal idEnum As Long) As String
    ' 1. Matriz temporal para recibir el diccionario
    Dim arrCampos As Variant
    
    ' 2. Obtenemos el diccionario exacto de la tabla solicitada
    arrCampos = TablasDB.CamposTablaDB(nombreTabla)
    
    ' 3. Prevención de Errores: Validamos que el Enum no exceda el tamaño de la matriz
    If idEnum >= LBound(arrCampos) And idEnum <= UBound(arrCampos) Then
        CampoDB = arrCampos(idEnum)
    Else
        Call App.SystemError("El Enum [" & idEnum & "] está fuera de rango para la tabla " & nombreTabla)
        CampoDB = ""
    End If
End Function
Public Sub FiltrarTabla(ByVal nombreTabla As String, ByVal nombreColumna As String, ByVal criterio As Variant, Optional ByVal busquedaParcial As Boolean = False)
    
    Dim tbl As ListObject
    Dim colIndex As Long
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(nombreTabla)
    
    On Error Resume Next
    Set tbl = ws.ListObjects(nombreTabla)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "Error de Arquitectura: La tabla '" & nombreTabla & "' no existe en esta hoja.", vbCritical, "Fallo de Filtro"
        Exit Sub
    End If
    
    If tbl.ShowAutoFilter Then
        If tbl.AutoFilter.FilterMode Then
            tbl.AutoFilter.ShowAllData
        End If
    End If
    
    On Error Resume Next
    colIndex = tbl.ListColumns(nombreColumna).Index
    On Error GoTo 0
    
    If colIndex = 0 Then
        MsgBox "Error de Tabla: La columna '" & nombreColumna & "' no se encontró.", vbExclamation
        Exit Sub
    End If
    
    If IsArray(criterio) Then
        tbl.Range.AutoFilter Field:=colIndex, Criteria1:=criterio, Operator:=xlFilterValues
        
    Else
        If busquedaParcial Then
            tbl.Range.AutoFilter Field:=colIndex, Criteria1:="*" & CStr(criterio) & "*"
        Else
            tbl.Range.AutoFilter Field:=colIndex, Criteria1:="=" & CStr(criterio)
        End If
    End If
    Application.Goto Reference:=tbl.HeaderRowRange(1, 1), Scroll:=True
    
End Sub
Function SetMantenimiento() As String
    Dim man As String
    While man = ""
        With TipoMantenimiento
            .Show
            man = .mantenimiento
        End With
        Unload TipoMantenimiento
    Wend
    SetMantenimiento = UCase(man)
End Function
