Attribute VB_Name = "DataBaseUtils"
Option Private Module

Global ActiveStocktech As Boolean 'Variable para inicializar la aplicación

Global GlobalTARGET As String 'Variable para almacenar el obveto al que se le hace algo en el AuditTrail
Global GlobalACTION As String 'Variable para almacenar la acción concreta para el AuditTrail
Global GlobalWHO As String 'Variable para almacenar a quien realiza la acción para el AuditTrail
Global GlobalCOMMIT As String 'Variable para guardar los detalles de una acción para el AuditTrail
Global GlobalLOG As Boolean 'Variable para verificación de Log en el AuditTrail
Global CURRENTUSER As String 'Variable para contener al usuario de inicio de sesión

' --- VARIABLES PRIVADAS DEL SINGLETON ---
Private GlobalDBConnection As ADODB.Connection
Private rs As ADODB.Recordset
Const MODULE_NAME As String = "DataBaseUtils"
Function resetGlobals()
    GlobalTARGET = Empty
    GlobalACTION = Empty
    GlobalWHO = Empty
    GlobalCOMMIT = Empty
    GlobalLOG = False
    USERLOGGED = Empty
    LEVELLOGGED = Empty
    LOGGED = Empty
    Seguridad.ClearCurrentUser
End Function


' Patrón de diseño SINGLETON para la conexión con la DB
Public Function GetDBConnection() As ADODB.Connection
    Dim rutaDB As String
    Dim passDb As String
    Dim strConn As String
    
    ' Si el Objeto no existe se crea
    If GlobalDBConnection Is Nothing Then
        Set GlobalDBConnection = New ADODB.Connection
    End If
    
    ' Si la conexión está cerrada se abre
    If GlobalDBConnection.State <> 1 Then
        
        ' Solo se utiliza recurso de CPU para desencriptar si se requiere
        rutaDB = DB()
        passDb = Seguridad.Desencriptar("9E67A3A5B7B5BF877F7397")
        
        strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & rutaDB & ";Jet OLEDB:Database Password=" & passDb & ";"
        
        GlobalDBConnection.ConnectionString = strConn
        GlobalDBConnection.Open
    End If
    
    ' Se retorna la conexión con la DB (objeto ADODB.connection)
    Set GetDBConnection = GlobalDBConnection
End Function

' Proceso para cerrar la conexión con la DB
Public Sub CloseDBConnection(Optional ByRef rs As ADODB.Recordset)
    ' 1. Cerrar y limpiar el Recordset (si se proporcionó)
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close ' 0 = adStateClosed
        Set rs = Nothing
    End If
    
    ' 2. Cerrar y limpiar la Conexión Global
    If Not GlobalDBConnection Is Nothing Then
        If GlobalDBConnection.State <> 0 Then GlobalDBConnection.Close
        Set GlobalDBConnection = Nothing
    End If
End Sub

Function AccionSQL(SQL As String) As Boolean
    Dim conn As ADODB.Connection
    
    ' =======================================================
    ' 1. ESCUDO DE INTERRUPCIÓN
    ' =======================================================
    ' Preparamos el motor para desviar cualquier fallo de conexión o sintaxis
    ' directamente a nuestro bloque de control de errores.
    On Error GoTo ErrorHandler
    
    ' =======================================================
    ' 2. APERTURA DE CONEXIÓN
    ' =======================================================
    ' El método GetDBConnection decide de forma inteligente si es necesario
    ' abrir una conexión nueva o reutilizar una existente en la memoria.
    Set conn = GetDBConnection()
    
    ' =======================================================
    ' 3. EJECUCIÓN DIRECTA
    ' =======================================================
    ' Disparamos el comando hacia el motor de base de datos.
    ' Al no esperar datos de regreso, no instanciamos ningún Recordset.
    conn.Execute SQL
    
    ' 4. Confirmación de Éxito
    AccionSQL = True
    Exit Function

ErrorHandler:
    ' =======================================================
    ' 5. MANEJO DE EXCEPCIONES UNIVERSAL
    ' =======================================================
    ' Evaluamos si el error proviene de una violación de regla de base de datos (Clave Duplicada)
    If InStr(1, LCase(Err.Description), "duplicad") > 0 Or InStr(1, LCase(Err.Description), "duplicate") > 0 Then
        
        ' Mensaje genérico aplicable a Equipos, SOLPEDs, Usuarios o Refacciones
        MsgBox "No se puede guardar el registro. El identificador principal ya existe en la base de datos.", vbExclamation, "Registro Duplicado"
        App.SystemError "Intento de registro duplicado abortado por el motor SQL."
        
    Else
        ' Fallos críticos (Pérdida de red, archivo de base de datos bloqueado o sintaxis SQL rota)
        MsgBox "Error de comunicación con la Base de Datos." & vbCrLf & _
               "Consulte a un administrador del sistema.", vbCritical, "Fallo Crítico"
               
        ' Trazabilidad exacta del fallo para el equipo de soporte
        App.SystemError "Error SQL: " & Err.Description & " | Comando: " & SQL
    End If
    
    ' 6. Confirmación de Fallo
    AccionSQL = False
End Function
' Función para obtener solo un dato de la DB
Function DatoDataBase(ByVal tabla As String, ByVal Filtro As String, ByVal valorBuscado As Variant, ByVal Resultado As String) As Variant
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim SQL As String
    Dim criterio As String
    
    ' 1. Casteo Dinámico para Access SQL
    criterio = ConsultasSQL.FormatoSQL(valorBuscado)
    
    ' 2. Construcción de la Consulta SQL
    SQL = "SELECT [" & Resultado & "] FROM [" & tabla & "] " & _
          "WHERE [" & Filtro & "] = " & CStr(criterio)
          
    On Error GoTo ManejadorErrores
    
    ' Se obtiene conexión
    Set conn = GetDBConnection()
    
    ' Se instancia el objeto donde se almacenará lo obtenido de la DB
    Set rs = New ADODB.Recordset
    
    ' 3 = adOpenStatic, 1 = adLockReadOnly
    rs.Open SQL, conn, 3, 1
    
    If Not rs.EOF Then
        DatoDataBase = rs.Fields(Resultado).Value ' Se obtiene solo el valor
    Else
        DatoDataBase = Empty
    End If
    
SalidaSegura:
    ' Solo se cierra el Recordset y se mantiene la conexión para mejorar la persistencia
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Function

ManejadorErrores:
    ' Tu lógica intacta
    MsgBox "Error con la obtención de información de la Base de Datos", vbCritical, "Error StockTech"
    Call App.SystemError("Error en DatoDataBase: " & Err.Description & vbNewLine & SQL)
    DatoDataBase = Empty
    Resume SalidaSegura
End Function

Function TableDataBase(ByVal tabla As String, _
                       ByVal columna As String, _
                       Optional conEncabezados As Boolean = False, _
                       Optional criterio As String = "") As Variant
    
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim SQL As String
    Dim vDatos As Variant, vFinal As Variant
    Dim i As Long, j As Long, numCols As Long, numFilas As Long

    On Error GoTo ErrorHandler

    Set conn = GetDBConnection()
    Set rs = New ADODB.Recordset
    
    ' Construcción dinámica de SQL
    SQL = "SELECT " & IIf(columna = "*", "*", "[" & columna & "]") & " FROM [" & tabla & "]"
    
    ' Inyección del filtro si el parámetro 'criterio' no está vacío
    If Trim(criterio) <> "" Then
        SQL = SQL & " WHERE " & criterio
    End If
    
    ' Abrimos Recordset: adOpenStatic (3), adLockReadOnly (1)
    rs.Open SQL, conn, 3, 1
    
    If Not rs.EOF Then
        numCols = rs.Fields.Count - 1
        
        If conEncabezados Then
            ' Capturamos nombres antes de que GetRows mueva el cursor
            ReDim vHeader(0 To numCols)
            For i = 0 To numCols
                vHeader(i) = rs.Fields(i).Name
            Next i
            
            vDatos = rs.GetRows
            numFilas = UBound(vDatos, 2)
            
            ' Redimensionamos vFinal (Columnas, Filas + 1 para el header)
            ReDim vFinal(0 To numCols, 0 To numFilas + 1)
            
            ' 1. Transferir encabezados
            For i = 0 To numCols
                vFinal(i, 0) = vHeader(i)
            Next i
            
            ' 2. Transferir datos
            For j = 0 To numFilas
                For i = 0 To numCols
                    vFinal(i, j + 1) = vDatos(i, j)
                Next i
            Next j
            
            TableDataBase = vFinal
        Else
            TableDataBase = rs.GetRows
        End If
    Else
        ' Manejo de resultados vacíos
        TableDataBase = Empty
    End If

CleanExit:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Function

ErrorHandler:
    ' Registro de errores profesional
    App.SystemError "Error en TableDataBase - Proyecto StockTech"
    App.SystemError "SQL: " & SQL
    App.SystemError "Descripción: " & Err.Description
    TableDataBase = Empty
    Resume CleanExit
End Function

Function QueryTableConsult(ByVal Tabla1 As String, ByVal Tabla2 As String, ByVal Relacion As String, _
                           ByVal CamposTabla2 As Variant, Optional ByVal CamposTabla1 As Variant, Optional ByVal Filtro As String = "", _
                           Optional ByVal valorBuscado As String = "") As Variant
    
    Dim conn As ADODB.Connection
    Dim SQL As String
    
    todosCampos = StringUtils.CrearCamposSQLTablas(CamposTabla2, CamposTabla1)
    
    ' 1. Armamos la base de la consulta (El JOIN principal)
    SQL = "SELECT " & todosCampos & " " & _
          "FROM [" & Tabla1 & "] AS U " & _
          "INNER JOIN [" & Tabla2 & "] AS P " & _
          "ON U.[" & Relacion & "] = P.[" & Relacion & "]"
          
    ' 2. Agregamos la cláusula WHERE solo si se proporcionaron los parámetros
    If Trim(Filtro) <> "" And Trim(valorBuscado) <> "" Then
        SQL = SQL & " WHERE U.[" & Filtro & "] = '" & valorBuscado & "'"
    End If

    On Error GoTo ErrorHandler
    
    ' Llamamos al Singleton
    Set conn = GetDBConnection()
    Set rs = New ADODB.Recordset
    
    ' Ejecutamos (0 = adOpenForwardOnly, 1 = adLockReadOnly)
    rs.Open SQL, conn, 3, 1
    
    If Not rs.EOF Then
        QueryTableConsult = rs.GetRows()
    Else
        QueryTableConsult = Empty
    End If

limpieza:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Function

ErrorHandler:
    MsgBox "Ha ocurrido un error. Consulte a un Administrador del Sistema", vbCritical
    App.SystemError ("Error en QueryTableConsult: " & Err.Description & vbCrLf & "SQL: " & SQL)
    QueryTableConsult = Empty
    Resume limpieza
End Function
Function AddNewUser(nombre As String, usuario As String, correo As String, nivel As String, contrasena As String) As Boolean
    Dim SQL As String
    Dim checkSQL As String
    Dim matrizAuditoria As Variant
    Dim usuariosEncontrados As Long
    
    ' =======================================================
    ' 1. AUDITORÍA PREVIA (Bloqueo de Duplicados)
    ' =======================================================
    checkSQL = "SELECT COUNT(*) FROM [Usuarios] WHERE [Usuario] = '" & usuario & "'"
    
    ' Utilizamos tu motor de extracción para traer el conteo
    matrizAuditoria = GetFromSQL(checkSQL)
    
    If IsArray(matrizAuditoria) Then
        usuariosEncontrados = matrizAuditoria(0, 0)
        
        ' Si el conteo es mayor a 0, el usuario ya existe. Abortamos operación.
        If usuariosEncontrados > 0 Then
            MsgBox "Imposible registrar: El identificador de usuario '" & usuario & "' ya está en uso.", vbExclamation, "Registro Duplicado"
            Exit Function
        End If
    End If
    
    ' =======================================================
    ' 2. INYECCIÓN DE DATOS (Inserción Segura)
    ' =======================================================
    Dim campos As Variant
    Dim valores As Variant
    
    campos = CamposTablaDB(TUsuarios)
    StringUtils.DropFromArray campos, 0
    
    valores = Array(nombre, usuario, correo, nivel, Seguridad.Encriptar(contrasena))
    
    If AddRegister(TUsuarios, campos, valores) Then
        MsgBox "Usuario creado con éxito.", vbInformation, "Gestión de Usuarios"
        AddNewUser = True
    End If
    
End Function

Function DeleteUser(usuario As String) As Boolean
    Dim SQL As String
    Dim respuesta As VbMsgBoxResult
    
    ' =======================================================
    ' PASO 1: ESCUDO DE AUTO-ELIMINACIÓN
    ' =======================================================
    ' Se compara el usuario que se desea eliminar con el usuario logueado en el sistema.
    ' Se utiliza UCase() en ambas partes para asegurar que no haya saltos de seguridad por mayúsculas.
    If UCase(usuario) = UCase(GlobalUsuarioActivo) Then
        MsgBox "Violación de seguridad: No puede eliminar su propio usuario mientras mantiene una sesión activa.", vbCritical, "Acción Denegada"
        Exit Function
    End If
    
    ' =======================================================
    ' PASO 2: BARRERA PSICOLÓGICA (Confirmación)
    ' =======================================================
    ' Al ser una acción destructiva física, se invoca un cuadro de diálogo
    ' que obliga al operador a confirmar la acción antes de tocar la base de datos.
    respuesta = MsgBox("¿Está seguro de  eliminar permanentemente al usuario '" & UCase(usuario) & "'?" & vbCrLf & _
                       "Esta acción es irreversible y borrará su acceso de inmediato.", vbYesNo + vbExclamation, "Confirmación de Eliminación Fïsica")
    
    ' Si el operador selecciona "No", la subrutina se aborta silenciosamente.
    If respuesta = vbNo Then
        Exit Function
    End If
    
    ' =======================================================
    ' PASO 3: ENSAMBLAJE DE LA INSTRUCCIÓN DE BORRADO
    ' =======================================================
    ' Se estructura la instrucción SQL DELETE FROM filtrando estrictamente
    ' por el identificador único del usuario para evitar daños colaterales.
    SQL = "DELETE FROM " & TUsuarios & " WHERE [" & CampoDB(TUsuarios, us_usuario) & "] = '" & usuario & "'"
    
    ' =======================================================
    ' PASO 4: EJECUCIÓN Y REGISTRO EN AUDITORÍA
    ' =======================================================
    ' Se envía la instrucción al motor principal. Si la base de datos lo elimina
    ' exitosamente, se registran los metadatos de la acción en el Audit Trail.
    If AccionSQL(SQL) Then
        DeleteUser = True
        MsgBox "El usuario ha sido eliminado permanentemente del sistema.", vbInformation, "Gestión de Usuarios"
    End If
End Function

Public Sub UserChanges(ByVal user As String, ByVal campo As CNameUsuarios, ByVal dato As Variant)
    Dim nombreColumna As String
    Dim valFinal As String
    Dim SQL As String
    
    ' 1. Traducción del Enum
    Select Case campo
        Case us_nombre: nombreColumna = "Nombre"
        Case us_usuario: nombreColumna = "Usuario"
        Case us_correo: nombreColumna = "Correo"
        Case us_nivel: nombreColumna = "Nivel"
        Case us_contrasenia: nombreColumna = "Contraseña"
        Case Else: Exit Sub
    End Select
    
    If campo = us_contrasenia Then dato = Seguridad.Encriptar(CStr(dato))
    
    ' 2. Llamamos al Motor Universal (UNA SOLA LÍNEA)
    valFinal = ConsultasSQL.FormatoSQL(dato)
    
    ' 3. Construcción del UPDATE
    SQL = "UPDATE [Usuarios] SET [" & nombreColumna & "] = " & valFinal & " " & _
          "WHERE [Usuario] = " & FormatoSQL(user) ' Reutilizamos la función para blindar el nombre de usuario
          
    If AccionSQL(SQL) Then
        MsgBox "El campo [" & nombreColumna & "] ha sido actualizado.", vbInformation, "Gestión"
    End If
End Sub
Public Function SystemLogDB(ByVal accion As String, ByVal objetivo As String) As Boolean
    Const PROC_NAME As String = "SystemLogDB"
    
    Dim tabla As String
    Dim campos As Variant
    Dim valores As Variant
    Dim comentarioUsuario As String
    
    tabla = TAuditTrail
        
    If DataBaseUtils.Auditoria("StockTech System", accion, objetivo, detalles) Then
        SystemLogDB = True
    End If
End Function
Public Function LogDB(ByVal accion As String, ByVal objetivo As String, ByVal detalles As String, Optional comentarios As Boolean = False) As Boolean
    Const PROC_NAME As String = "LogDB"
    
    Dim tabla As String
    Dim campos As Variant
    Dim valores As Variant
    Dim comentarioUsuario As String
    
    tabla = TAuditTrail
        
    If DataBaseUtils.Auditoria(Environ("USERNAME"), accion, objetivo, detalles, comentarios) Then
        LogDB = True
    End If
End Function
Public Function Auditoria(ByVal usuario As String, ByVal accion As String, ByVal objetivo As String, ByVal detalles As String, Optional ByVal WithCommit As Boolean = False) As Boolean
    Const PROC_NAME As String = "Auditoria"
    Dim campos As Variant
    Dim valores As Variant
    Dim comentarioFinal As String
    
    On Error GoTo ErrorHandler
    
    comentarioFinal = "N/A" ' Valor por defecto

    If WithCommit Then
        ' 1. Llamada encapsulada
        With AuditTrail
            .AccionLabel.Caption = accion
            .Resultado = "" ' Limpiamos rastro anterior
            .Show ' El código se pausa aquí (Modal)
            
            ' 2. Al regresar (tras el .Hide), leemos la propiedad pública
            comentarioFinal = .Resultado
        End With
        
        ' 3. Validación de salida
        If Trim(comentarioFinal) = "" Or comentarioFinal = "N/A" Then
            MsgBox "Acción cancelada: Justificación obligatoria.", vbExclamation, "StockTech"
            Unload AuditTrail ' Limpiamos memoria ahora sí
            Auditoria = False
            Exit Function
        End If
        
        Unload AuditTrail ' Descargamos tras leer exitosamente
    End If
    
    ' 4. Preparación y Envío (Usamos comentarioFinal, no la global Commit)
    Dim CamposAudit As Variant
    CamposAudit = CamposTablaDB(TAuditTrail)
    
    campos = Array(CamposAudit(aud_usuario), CamposAudit(aud_fecha), CamposAudit(aud_accion), _
                    CamposAudit(aud_objetivo), CamposAudit(aud_detalles), CamposAudit(aud_comentario))
                    
    valores = Array(usuario, Now(), accion, objetivo, detalles, comentarioFinal)
    
    If DataBaseUtils.AddRegister(TablasDB.TAuditTrail, campos, valores) Then
        Auditoria = True
    Else
        Err.Raise vbObjectError + 516, , "Error al escribir en TAuditTrail."
    End If
    
    Exit Function
    
ErrorHandler:
    App.SystemError "Fallo en [" & MODULE_NAME & "." & PROC_NAME & "] - " & Err.Description
    Auditoria = False
End Function

Public Function DB() As String
    Dim rutaDefault As String
    Dim rutaEncriptada As String
    Dim rutaFallback As String
    
    ' 1. Prioridad Absoluta: Buscar junto al archivo .xlam
    rutaDefault = ThisWorkbook.path & "\StockTech DB.accdb"
    
    ' Evaluamos si la base de datos existe físicamente en esa ruta local
    If Dir(rutaDefault) <> "" Then
        DB = rutaDefault
        Exit Function
    End If
    
    ' 2. Plan de Contingencia (Fallback): Leer el Registro de Windows
    rutaEncriptada = GetSetting("StockTech", "Conexion", "RutaDB", "")
    
    If rutaEncriptada <> "" Then
        ' Desencriptamos la ruta almacenada históricamente
        rutaFallback = Seguridad.Desencriptar(rutaEncriptada)
        
        ' Validamos que el archivo en la ruta del registro realmente siga existiendo
        If Dir(rutaFallback) <> "" Then
            DB = rutaFallback
            Exit Function
        End If
    End If
    
    ' 3. Estado de Falla Crítica
    ' Si la base de datos fue eliminada, movida o no está conectada
    Call App.SystemError("No se localizó la base de datos ni en la raíz del sistema ni en los registros guardados.")
    DB = ""
End Function
Public Function UniqueValues(ByVal columna As Variant, ByVal origenDatos As Variant) As Variant
    Const PROC_NAME As String = "UniqueValues"
    On Error GoTo ErrorHandler
    
    ' =====================================================================
    ' RUTA A: PROCESAMIENTO EN BASE DE DATOS (El origen es un String)
    ' =====================================================================
    If VarType(origenDatos) = vbString Then
        Dim tabla As String
        Dim conn As ADODB.Connection
        Dim rs As ADODB.Recordset
        Dim SQL As String
        Dim resultadoDB As Variant
        Dim lista() As Variant
        Dim Filas As Long, i As Long
        
        tabla = CStr(origenDatos)
        
        ' 1. Traducción Automática con Escudo de Límites
        If IsNumeric(columna) Then
            Dim arrCampos As Variant
            arrCampos = TablasDB.CamposTablaDB(tabla)
            
            If columna >= LBound(arrCampos) And columna <= UBound(arrCampos) Then
                columna = arrCampos(columna)
            Else
                Call App.SystemError("Fallo de Arquitectura: El Enum (" & columna & ") supera las columnas de [" & tabla & "].")
                UniqueValues = Empty
                Exit Function
            End If
        End If
        
        ' 2. Petición SQL optimizada
        SQL = "SELECT DISTINCT [" & columna & "] FROM [" & tabla & "] WHERE [" & columna & "] IS NOT NULL"
              
        ' 3. Ejecución
        Set conn = DataBaseUtils.GetDBConnection()
        Set rs = New ADODB.Recordset
        rs.Open SQL, conn, 0, 1
            
        ' 4. Extracción
        If Not rs.EOF Then
            resultadoDB = rs.GetRows()
            Filas = UBound(resultadoDB, 2)
            ReDim lista(0 To Filas)
            For i = 0 To Filas
                lista(i) = resultadoDB(0, i) & ""
            Next i
            UniqueValues = lista
        Else
            UniqueValues = Empty
        End If
        
        ' Limpieza ADO
        On Error Resume Next
        If Not rs Is Nothing Then
            If rs.State = 1 Then rs.Close
            Set rs = Nothing
        End If
        Exit Function

    ' =====================================================================
    ' RUTA B: PROCESAMIENTO EN MEMORIA RAM (El origen es una Matriz GetRows)
    ' =====================================================================
    ElseIf IsArray(origenDatos) Then
        Dim arrDatos As Variant
        Dim dictUnicos As Object
        Dim valor As String
        Dim f As Long
        
        arrDatos = origenDatos
        Set dictUnicos = CreateObject("Scripting.Dictionary")
        dictUnicos.CompareMode = 1 ' vbTextCompare (Insensible a mayúsculas para evitar duplicados como "A" y "a")
        
        ' Validación de integridad
        If Not IsNumeric(columna) Then
            Call App.SystemError("Fallo de Arquitectura: Para matrices en memoria, la columna debe ser un índice numérico.")
            UniqueValues = Empty
            Exit Function
        End If
        
        If columna < LBound(arrDatos, 1) Or columna > UBound(arrDatos, 1) Then
            Call App.SystemError("Fallo de Arquitectura: El índice (" & columna & ") está fuera de la matriz de datos.")
            UniqueValues = Empty
            Exit Function
        End If
        
        ' Recorremos la matriz (GetRows es bidimensional: 0 To Columnas, 0 To Filas)
        For f = LBound(arrDatos, 2) To UBound(arrDatos, 2)
            If Not IsNull(arrDatos(columna, f)) Then
                valor = Trim(CStr(arrDatos(columna, f)))
                ' Solo agregamos si no está vacío y si no existe previamente en el diccionario
                If valor <> "" And Not dictUnicos.Exists(valor) Then
                    dictUnicos.Add valor, valor
                End If
            End If
        Next f
        
        ' Devolvemos las llaves del diccionario como arreglo unidimensional
        If dictUnicos.Count > 0 Then
            UniqueValues = dictUnicos.Keys()
        Else
            UniqueValues = Empty
        End If
        
        Exit Function

    ' =====================================================================
    ' RUTA C: EXCEPCIÓN DE TIPO DE DATO
    ' =====================================================================
    Else
        Call App.SystemError("Fallo de Arquitectura: 'origenDatos' debe ser el nombre de una tabla o una matriz.")
        UniqueValues = Empty
        Exit Function
    End If

ErrorHandler:
    Call App.SystemError("Error en [" & PROC_NAME & "]: " & Err.Description)
    UniqueValues = Empty
    ' Aseguramos que los objetos ADO se cierren si falló en la Ruta A
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
    End If
End Function
Public Function UpdateTable(ByVal tabla As String, Optional Silencio As Boolean = False) As String
    Const PROC_NAME As String = "UpdateTable"
    
    Dim con As Object, rs As Object
    Dim hoja As Worksheet, SelectedTable As ListObject
    Dim columna As ListColumn
    Dim filaTable As Long
    Dim idBusqueda As Variant
    Dim strModificados As String, strNuevos As String
    Dim cambios As Long, addedElement As Long
    Dim ColumnaIdName As String
    Dim filaNecesitaUpdate As Boolean
    Dim valCelda As Variant
    
    ' =======================================================
    ' 1. ESCUDO DE INTERRUPCIÓN
    ' =======================================================
    On Error GoTo ErrorHandler
    
    Set con = DataBaseUtils.GetDBConnection()
    Set hoja = ActiveWorkbook.Worksheets(tabla)
    Set SelectedTable = hoja.ListObjects(1)
    
    Set rs = CreateObject("ADODB.Recordset")
    ' 1 = adOpenKeyset, 3 = adLockOptimistic
    rs.Open "SELECT * FROM [" & tabla & "]", con, 1, 3
    
    ColumnaIdName = rs.Fields(0).Name
    strModificados = ""
    strNuevos = ""

    ' =======================================================
    ' 2. BUCLE PRINCIPAL DE SINCRONIZACIÓN
    ' =======================================================
    For filaTable = 1 To SelectedTable.ListRows.Count
        filaNecesitaUpdate = False
        idBusqueda = SelectedTable.ListColumns(1).DataBodyRange(filaTable).Value
        
        ' --- LA CORRECCIÓN CRÍTICA (Blindaje contra Error 13) ---
        ' Usamos Trim y Val para forzar la conversión matemática segura
        If Trim(CStr(idBusqueda)) <> "" And Val(idBusqueda) <> 0 Then
            rs.Filter = "[" & ColumnaIdName & "] = " & CLng(idBusqueda)
        Else
            ' Si la celda está vacía o tiene letras, forzamos el EOF (-1)
            rs.Filter = "[" & ColumnaIdName & "] = -1"
        End If

        If Not rs.EOF Then
            ' =======================================================
            ' A. LÓGICA DE ACTUALIZACIÓN
            ' =======================================================
            For Each columna In SelectedTable.ListColumns
                If columna.Index <> 1 Then
                    valCelda = columna.DataBodyRange(filaTable).Value
                    
                    ' Validamos si hubo un cambio real
                    If CStr(rs.Fields(columna.Name).Value & "") <> CStr(valCelda & "") Then
                        ' Inyección Segura
                        If Trim(CStr(valCelda)) = "" Then
                            rs.Fields(columna.Name).Value = Null
                        Else
                            rs.Fields(columna.Name).Value = valCelda
                        End If
                        filaNecesitaUpdate = True
                    End If
                End If
            Next
            
            If filaNecesitaUpdate Then
                rs.Update
                cambios = cambios + 1
                strModificados = strModificados & IIf(strModificados = "", "", ", ") & idBusqueda
            End If
            
        Else
            ' =======================================================
            ' B. LÓGICA DE ADICIÓN (Altas Nuevas)
            ' =======================================================
            rs.Filter = 0 ' Limpiamos el filtro estrictamente antes de agregar
            rs.AddNew
            
            For Each columna In SelectedTable.ListColumns
                If columna.Index <> 1 Then
                    valCelda = columna.DataBodyRange(filaTable).Value
                    
                    ' Inyección Segura: Convertimos celdas vacías a Null SQL
                    If Trim(CStr(valCelda)) = "" Then
                        rs.Fields(columna.Name).Value = Null
                    Else
                        rs.Fields(columna.Name).Value = valCelda
                    End If
                End If
            Next columna
            
            rs.Update
            addedElement = addedElement + 1
            
            ' Captura del ID generado por Access
            Dim nuevoID As Long
            nuevoID = rs.Fields(ColumnaIdName).Value
            strNuevos = strNuevos & IIf(strNuevos = "", "", ", ") & nuevoID
        End If
        
        ' Limpieza del filtro para la siguiente fila
        rs.Filter = 0
    Next

    ' =======================================================
    ' 3. REPORTE Y TRAZABILIDAD (Audit Trail)
    ' =======================================================
    Dim reporteFinal As String
    reporteFinal = ""
    
    If strModificados <> "" Then reporteFinal = "IDs MODIFICADOS: [" & strModificados & "]"
    If strNuevos <> "" Then
        reporteFinal = reporteFinal & IIf(reporteFinal = "", "", " | ") & "IDs NUEVOS: [" & strNuevos & "]"
    End If
    
    UpdateTable = IIf(reporteFinal = "", "SIN CAMBIOS", reporteFinal)

    ' =======================================================
    ' 4. NOTIFICACIÓN Y RECARGA
    ' =======================================================
    If cambios > 0 Or addedElement > 0 Then
        If Not Silencio Then
            MsgBox "Sincronización de la tabla '" & tabla & "' completada:" & vbCrLf & _
                   "- Registros modificados: " & cambios & vbCrLf & _
                   "- Registros nuevos: " & addedElement, vbInformation, "StockTech Sync"
        End If
        ' Refrescamos la tabla para mostrar los Autonuméricos generados
        Call DataBaseUtils.GetExcelTable(tabla, refresh:=True)
    End If

limpieza:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Application.ScreenUpdating = True
    Exit Function

ErrorHandler:
    ' Modificado para mostrar exactamente por qué falló si vuelve a ocurrir
    App.SystemError "Excepción crítica en UpdateTable [" & tabla & "] - Error de Ejecución: " & Err.Description
    UpdateTable = "ERROR"
    Resume limpieza
End Function
Public Function GetExcelTable(ByVal nombreTabla As String, _
                              Optional ByVal vDataIn As Variant, _
                              Optional ByVal vHeaders As Variant, _
                              Optional ByVal refresh As Boolean = False) As Boolean
    
    Const PROC_NAME As String = "GetExcelTable"
    Dim vData As Variant
    Dim vEncabezados As Variant
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim c As Long, f As Long
    Dim esReportePersonalizado As Boolean
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ' 1. Origen de Datos y Encabezados
    If IsMissing(vDataIn) Or IsEmpty(vDataIn) Then
        vData = DataBaseUtils.TableDataBase(nombreTabla, "*", False)
        vEncabezados = TablasDB.CamposTablaDB(nombreTabla)
        esReportePersonalizado = False
    Else
        vData = vDataIn
        vEncabezados = vHeaders
        esReportePersonalizado = True
    End If
    
    If IsEmpty(vData) Then
        MsgBox "No hay datos para mostrar en '" & nombreTabla & "'", vbExclamation
        GetExcelTable = False
        Exit Function
    End If

    ' 2. Evaluación de Hoja Existente
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(nombreTabla)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        If Not refresh Then
            If MsgBox("La hoja '" & nombreTabla & "' ya existe. ¿Desea sobrescribirla?", vbYesNo + vbQuestion) = vbNo Then
                ws.Select
                GetExcelTable = True
                Exit Function
            End If
            
            If Not esReportePersonalizado Then
                If MsgBox("¿Desea sincronizar cambios antes de refrescar?", vbYesNo) = vbYes Then
                    Call DataBaseUtils.UpdateTable(nombreTabla)
                End If
            End If
        End If
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    ' 3. Creación de la Hoja
    Set ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    ws.Name = nombreTabla
    Call Seguridad.MarkStockTechSheet(ws)

    ' 4. Escritura de Encabezados (Fila 1)
    If Not IsEmpty(vEncabezados) Then
        For c = 0 To UBound(vEncabezados)
            ws.Cells(1, c + 1).Value = vEncabezados(c)
        Next c
    End If
    
    ' 5. Escritura de Datos (Iniciamos en Fila 2)
    For f = 0 To UBound(vData, 2)
        For c = 0 To UBound(vData, 1)
            ws.Cells(f + 2, c + 1).Value = vData(c, f)
        Next c
    Next f

    ' 6. Conversión a ListObject (Tabla oficial de Excel)
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
    lo.Name = Replace(nombreTabla, " ", "_")
    lo.TableStyle = "TableStyleMedium2"

    ' =======================================================
    ' 6.1 ORDENAMIENTO ASCENDENTE DE LA TABLA
    ' =======================================================
    With lo.Sort
        ' Paso A: Limpiar cualquier criterio de ordenamiento residual en la memoria de la tabla
        .SortFields.Clear
        
        ' Paso B: Configurar la llave de ordenamiento apuntando estrictamente a la Columna 1
        .SortFields.Add Key:=lo.ListColumns(1).Range, _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        
        ' Paso C: Indicar que la tabla tiene encabezados para evitar que la fila 1 se mezcle
        .Header = xlYes
        
        ' Paso D: Ejecutar el motor de ordenamiento físico
        .Apply
    End With
    ' =======================================================

    ' 7. Formato Final
    ws.Columns.AutoFit
    If Not esReportePersonalizado Then
        ws.Rows(1).Locked = True
        lo.ListColumns(CampoDB(nombreTabla, 0)).Range.EntireColumn.Hidden = True
    End If
    
    
    Application.ScreenUpdating = True
    GetExcelTable = True
    App.EnfocarStockTechLabels
    Exit Function

ErrorHandler:
    App.SystemError Err.Description
    Application.ScreenUpdating = True
    GetExcelTable = False
End Function

Public Function AddRegister(ByVal tabla As String, ByVal campos As Variant, ByVal valores As Variant) As Boolean
    Dim conn As Object
    Dim SQL As String
    Dim strCampos As String
    Dim strValores As String
    Dim i As Long
    Dim valFinal As String

    ' =======================================================
    ' 0. NORMALIZACIÓN DE ENTRADAS (Casting a Matrices)
    ' =======================================================
    If Not IsArray(campos) Then campos = Array(campos)
    If Not IsArray(valores) Then valores = Array(valores)

    ' =======================================================
    ' 1. VALIDACIÓN DE PARIDAD
    ' =======================================================
    If UBound(campos) <> UBound(valores) Then
        MsgBox "Error de consistencia: Campos (" & UBound(campos) + 1 & ") vs Valores (" & UBound(valores) + 1 & ")", vbCritical
        AddRegister = False
        Exit Function
    End If

    ' =======================================================
    ' 2. ENSAMBLAJE DINÁMICO
    ' =======================================================
    For i = LBound(campos) To UBound(campos)
        strCampos = strCampos & "[" & campos(i) & "], "
        valFinal = ConsultasSQL.FormatoSQL(valores(i))
        strValores = strValores & valFinal & ", "
    Next i

    ' =======================================================
    ' 3. LIMPIEZA Y EJECUCIÓN
    ' =======================================================
    strCampos = Left(strCampos, Len(strCampos) - 2)
    strValores = Left(strValores, Len(strValores) - 2)

    SQL = "INSERT INTO [" & tabla & "] (" & strCampos & ") VALUES (" & strValores & ")"

    On Error GoTo ErrorHandler
    Set conn = GetDBConnection()
    
    If conn.State = 0 Then conn.Open
    
    conn.Execute SQL
    AddRegister = True

limpieza:
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Exit Function

ErrorHandler:
    App.SystemError "Error en inserción DB: " & Err.Description & vbCrLf & "SQL Generado: " & SQL
    AddRegister = False
    Resume limpieza
End Function
Public Function SetRegister(ByVal tabla As String, _
                            ByVal camposUpdate As Variant, ByVal valoresUpdate As Variant, _
                            ByVal camposWhere As Variant, ByVal valoresWhere As Variant) As Boolean
                            
    Dim conn As Object
    Dim SQL As String
    Dim strSet As String
    Dim strWhere As String
    Dim i As Long
    
    ' =======================================================
    ' 0. NORMALIZACIÓN DE ENTRADAS (Casting a Matrices)
    ' =======================================================
    If Not IsArray(camposUpdate) Then camposUpdate = Array(camposUpdate)
    If Not IsArray(valoresUpdate) Then valoresUpdate = Array(valoresUpdate)
    If Not IsArray(camposWhere) Then camposWhere = Array(camposWhere)
    If Not IsArray(valoresWhere) Then valoresWhere = Array(valoresWhere)
    
    ' =======================================================
    ' 1. VALIDACIÓN DE PARIDAD (Seguridad de Matrices)
    ' =======================================================
    If UBound(camposUpdate) <> UBound(valoresUpdate) Then
        MsgBox "Error SET: La cantidad de campos a actualizar no coincide con los valores proporcionados.", vbCritical, "SetRegister"
        SetRegister = False
        Exit Function
    End If
    
    If UBound(camposWhere) <> UBound(valoresWhere) Then
        MsgBox "Error WHERE: La cantidad de campos de filtro no coincide con los valores proporcionados.", vbCritical, "SetRegister"
        SetRegister = False
        Exit Function
    End If

    ' =======================================================
    ' 2. CONSTRUCCIÓN DE LA CLÁUSULA "SET"
    ' =======================================================
    For i = LBound(camposUpdate) To UBound(camposUpdate)
        strSet = strSet & "[" & camposUpdate(i) & "] = " & ConsultasSQL.FormatoSQL(valoresUpdate(i)) & ", "
    Next i
    strSet = Left(strSet, Len(strSet) - 2)

    ' =======================================================
    ' 3. CONSTRUCCIÓN DE LA CLÁUSULA "WHERE"
    ' =======================================================
    For i = LBound(camposWhere) To UBound(camposWhere)
        strWhere = strWhere & "[" & camposWhere(i) & "] = " & ConsultasSQL.FormatoSQL(valoresWhere(i)) & " AND "
    Next i
    strWhere = Left(strWhere, Len(strWhere) - 5)

    ' =======================================================
    ' 4. ENSAMBLAJE SQL Y EJECUCIÓN
    ' =======================================================
    SQL = "UPDATE [" & tabla & "] SET " & strSet & " WHERE " & strWhere

    On Error GoTo ErrorHandler
    Set conn = GetDBConnection()
    
    If conn.State = 0 Then conn.Open
    conn.Execute SQL
    SetRegister = True

limpieza:
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Exit Function

ErrorHandler:
    App.SystemError "Error en actualización DB: " & Err.Description & vbCrLf & "SQL Generado: " & SQL
    SetRegister = False
    Resume limpieza
End Function
Function ConsultaSQL(SQL As String, con As ADODB.Connection) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    
    On Error GoTo ErrorHandler

    rs.Open SQL, con, 3, 1
    
    
    Set ConsultaSQL = rs
    Exit Function

ErrorHandler:
    MsgBox "Error en la consulta: " & Err.Description, vbCritical
    
    Set ConsultaSQL = Nothing
End Function
Public Sub CloseRS(ByRef rs As ADODB.Recordset)
    On Error Resume Next ' Evita que el código se detenga si el RS ya estaba muerto
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
        Set rs = Nothing
    End If
    On Error GoTo 0
End Sub
Public Function GetColumnName(ByVal tabla As String, ByVal indiceColumna As Integer) As String
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim SQL As String
    
    ' Mantenemos TOP 1 para que el motor SQL no trabaje de más leyendo datos
    SQL = "SELECT TOP 1 * FROM [" & tabla & "]"
    
    On Error GoTo ManejadorErrores
    
    Set conn = GetDBConnection()
    Set rs = New ADODB.Recordset
    rs.Open SQL, conn, adOpenStatic, adLockReadOnly
    
    ' --- VALIDACIÓN DE INTEGRIDAD (Fuera de Rango) ---
    If rs.Fields.Count > 0 Then
        ' Verificamos que el índice solicitado realmente exista en la tabla
        If indiceColumna >= 0 And indiceColumna < rs.Fields.Count Then
            GetColumnName = rs.Fields(indiceColumna).Name
        Else
            MsgBox "Error de rango: La tabla '" & tabla & "' tiene " & rs.Fields.Count & _
                   " columnas. (Índices válidos del 0 al " & rs.Fields.Count - 1 & ").", _
                   vbExclamation, "StockTech"
            GetColumnName = ""
        End If
    Else
        GetColumnName = ""
    End If
    On Error GoTo 0
SalidaSegura:
    Call CloseRS(rs)
    Exit Function

ManejadorErrores:
    MsgBox "No se pudo obtener el esquema de la tabla: " & tabla, vbCritical
    
    App.SystemError ("Error con la conexión o el SQL: " & SQL)
    ObtenerNombreColumna = ""
    Resume SalidaSegura
End Function
Public Function DeleteRegisterByID(ByVal tabla As String, ByVal valorID As Variant) As Boolean
    Const PROC_NAME As String = "DeleteRegisterByID"
    Dim conn As Object
    Dim SQL As String
    Dim valFinal As String
    Dim registrosAfectados As Long
    Dim columnID As String
    
    columnID = GetColumnName(tabla, 0)
    
    ' 1. Casteo del valor ID
    Select Case VarType(valorID)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            valFinal = CStr(valorID)
        Case vbString
            If Len(Trim(valorID)) = 0 Then
                MsgBox "El ID no puede estar vacío.", vbCritical, "Error de Integridad"
                DeleteRegisterByID = False
                Exit Function
            End If
            valFinal = "'" & Replace(valorID, "'", "''") & "'"
        Case Else
            MsgBox "El tipo de dato proporcionado para el ID no es válido para eliminación.", vbCritical, "Error de Integridad"
            DeleteRegisterByID = False
            Exit Function
    End Select

    ' 2. Construcción de la sentencia DELETE (Corregido a columnID)
    SQL = "DELETE FROM [" & tabla & "] WHERE [" & columnID & "] = " & valFinal

    On Error GoTo ErrorHandler
    Set conn = GetDBConnection()
    If conn.State = 0 Then conn.Open
    
    ' 3. Ejecución
    conn.Execute SQL, registrosAfectados
    
    If registrosAfectados > 0 Then
        DeleteRegisterByID = True
    Else
        DeleteRegisterByID = False
    End If

limpieza:
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Exit Function

ErrorHandler:
    App.SystemError "Error en eliminación DB: " & Err.Description & vbCrLf & "SQL Generado: " & SQL
    DeleteRegisterByID = False
    Resume limpieza
End Function
Public Function SetDataByID(ByVal tabla As String, ByVal campos As Variant, ByVal valores As Variant, ByVal valorID As Variant) As Boolean
    Const PROC_NAME As String = "SetDataByID"
    Dim conn As Object
    Dim SQL As String
    Dim strSet As String
    Dim i As Long
    Dim valFinal As String
    Dim idFinal As String
    Dim registrosAfectados As Long
    Dim columnID As String
    
    columnID = GetColumnName(tabla, 0)

    If Not IsArray(campos) Then campos = Array(campos)
    If Not IsArray(valores) Then valores = Array(valores)

    If UBound(campos) <> UBound(valores) Then
        MsgBox "Error de consistencia: Campos (" & UBound(campos) + 1 & ") vs Valores (" & UBound(valores) + 1 & ")", vbCritical, "Error StockTech"
        SetDataByID = False ' Corregido el nombre de retorno
        Exit Function
    End If

    For i = LBound(campos) To UBound(campos)
        valFinal = ConsultasSQL.FormatoSQL(valores(i))
        
        strSet = strSet & "[" & campos(i) & "] = " & valFinal & ", "
    Next i

    strSet = Left(strSet, Len(strSet) - 2)

    Select Case VarType(valorID)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            idFinal = CStr(valorID)
        Case vbString
            If Len(Trim(valorID)) = 0 Then
                MsgBox "El ID de actualización no puede estar vacío.", vbCritical, "Error de Integridad"
                SetDataByID = False ' Corregido el nombre de retorno
                Exit Function
            End If
            idFinal = "'" & Replace(valorID, "'", "''") & "'"
        Case Else
            MsgBox "Tipo de ID inválido para actualización.", vbCritical
            SetDataByID = False ' Corregido el nombre de retorno
            Exit Function
    End Select

    ' 4. Ensamblaje de la Sentencia SQL Final (Corregido a columnID)
    SQL = "UPDATE [" & tabla & "] SET " & strSet & " WHERE [" & columnID & "] = " & idFinal

    On Error GoTo ErrorHandler
    Set conn = GetDBConnection()
    If conn.State = 0 Then conn.Open
    
    conn.Execute SQL, registrosAfectados
    
    If registrosAfectados > 0 Then
        SetDataByID = True ' Corregido el nombre de retorno
    Else
        SetDataByID = False ' Corregido el nombre de retorno
    End If

limpieza:
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Exit Function

ErrorHandler:
    App.SystemError "Error en actualización DB: " & Err.Description & vbCrLf & "SQL Generado: " & SQL
    SetDataByID = False ' Corregido el nombre de retorno
    Resume limpieza
End Function
Public Function GetLastIdFromDB(ByVal nombreTabla As String, ByVal nombreColumnaID As String) As Long
    Const PROC_NAME As String = "GetLastIdFromDB"
    Dim rs As Object
    Dim SQL As String
    
    On Error GoTo ErrorHandler
    
    SQL = "SELECT IIf(IsNull(MAX([" & nombreColumnaID & "])), 0, MAX([" & nombreColumnaID & "])) " & _
          "FROM [" & nombreTabla & "]"
    
    ' Ejecutamos la consulta
    Set rs = DataBaseUtils.ConsultaSQL(SQL, DataBaseUtils.GetDBConnection)
    
    If Not rs.EOF Then
        GetLastIdFromDB = CLng(rs.Fields(0).Value)
    Else
        GetLastIdFromDB = 0
    End If

limpieza:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Function

ErrorHandler:
    ' Registro en la Caja Negra y Audit Trail
    App.SystemError "ERROR REGISTRADO: Al consultar tabla [" & nombreTabla & "] - " & Err.Description
    GetLastIdFromDB = -1 ' Devolvemos -1 para indicar un error técnico de conexión
    Resume limpieza
End Function
Public Function GetFromSQL(SQL As String, Optional WithHeaders As Boolean = False) As Variant
    Const PROC_NAME As String = "GetFromSQL"
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Dim vDatos As Variant, vFinal As Variant
    Dim i As Long, j As Long, numCols As Long, numFilas As Long

    On Error GoTo ErrorHandler

    Set conn = GetDBConnection()
    Set rs = New ADODB.Recordset
    
    rs.Open SQL, conn, 3, 1
    
    If Not rs.EOF Then
        numCols = rs.Fields.Count - 1
        
        If WithHeaders Then
            
            ReDim vHeader(0 To numCols)
            For i = 0 To numCols
                vHeader(i) = rs.Fields(i).Name
            Next i
            
            vDatos = rs.GetRows
            numFilas = UBound(vDatos, 2)
            
            ' Redimensionamos vFinal (Columnas, Filas + 1 para el header)
            ReDim vFinal(0 To numCols, 0 To numFilas + 1)
            
            ' 1. Transferir encabezados
            For i = 0 To numCols
                vFinal(i, 0) = vHeader(i)
            Next i
            
            ' 2. Transferir datos
            For j = 0 To numFilas
                For i = 0 To numCols
                    vFinal(i, j + 1) = vDatos(i, j)
                Next i
            Next j
            
            GetFromSQL = vFinal
        Else
            GetFromSQL = rs.GetRows
        End If
    Else
        ' Manejo de resultados vacíos
        GetFromSQL = Empty
    End If

CleanExit:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Function

ErrorHandler:
    ' Registro de errores profesional
    App.SystemError "Error en GetFromSQL - Proyecto StockTech"
    App.SystemError "SQL: " & SQL
    App.SystemError "Descripción: " & Err.Description
    GetFromSQL = Empty
    Resume CleanExit
End Function
Public Function RollBack(tabla As String, Optional Silencioso As Boolean = False) As Boolean
    Const PROC_NAME As String = "RollBack"
    Dim UltimoID As Integer
    UltimoID = DataBaseUtils.GetLastIdFromDB(tabla, CampoDB(tabla, 0))
    RollBack = DataBaseUtils.DeleteRegisterByID(tabla, UltimoID)
    If Not Silencioso Then
        App.SystemError "Se realiza ROLLBACK en la tabla " & tabla
    End If
End Function
