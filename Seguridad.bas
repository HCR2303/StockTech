Attribute VB_Name = "Seguridad"
Option Private Module
Option Explicit
Private Const SECRETO As String = "LA_FORMULA_SECRETA_DE_LA_VIDA_ES_LA_PA-CIENCIA"
Public HoraCierre As Double

' Variables para almacenar las contraseñas desencriptadas en la RAM
Private CachePassLibro As String
Private CachePassHoja As String
Const MODULE_NAME As String = "Seguridad"
Public LibroStockTech As Workbook ' Variable que recordará cuál es el libro oficial de esta sesión
Public Function Encriptar(ByVal texto As String) As String
    Dim i As Long
    Dim charTexto As Integer
    Dim charClave As Integer
    Dim cifrado As String
    Dim resultadoHex As String
    
    If texto = "" Then Exit Function
    
    
    For i = 1 To Len(texto)
        charTexto = Asc(Mid(texto, i, 1))
        
        charClave = Asc(Mid(SECRETO, ((i - 1) Mod Len(SECRETO)) + 1, 1))
        
        cifrado = cifrado & Chr((charTexto + charClave) Mod 255)
    Next i
    
    For i = 1 To Len(cifrado)
        resultadoHex = resultadoHex & Right("0" & Hex(Asc(Mid(cifrado, i, 1))), 2)
    Next i
    
    Encriptar = resultadoHex
End Function
Public Function Desencriptar(ByVal textoHex As String) As String
    Dim i As Long
    Dim charOriginal As Integer
    Dim charClave As Integer
    Dim textoIntermedio As String
    Dim Resultado As String
    Dim hexPar As String
    
    If textoHex = "" Then Exit Function
    
    For i = 1 To Len(textoHex) Step 2
        hexPar = Mid(textoHex, i, 2)
        textoIntermedio = textoIntermedio & Chr("&H" & hexPar)
    Next i
    
    
    For i = 1 To Len(textoIntermedio)
        charOriginal = Asc(Mid(textoIntermedio, i, 1))
        charClave = Asc(Mid(SECRETO, ((i - 1) Mod Len(SECRETO)) + 1, 1))
        
        Resultado = Resultado & Chr((charOriginal - charClave + 255) Mod 255)
    Next i
    
    Desencriptar = Resultado
End Function
Private Function GetPassLibro() As String
    ' Si la variable está vacía, consultamos a la DB
    If CachePassLibro = "" Then
        Dim passDb As Variant
        passDb = DataBaseUtils.DatoDataBase(TAccesos, CampoDB(TAccesos, acc_tipo), "Libro", CampoDB(TAccesos, acc_contrasenia))
        If Not IsEmpty(passDb) And Not IsNull(passDb) Then
            CachePassLibro = Desencriptar(CStr(passDb))
        End If
    End If
    GetPassLibro = CachePassLibro
End Function

Private Function GetPassHoja() As String
    ' Si la variable está vacía, consultamos a la DB
    If CachePassHoja = "" Then
        Dim passDb As Variant
        passDb = DataBaseUtils.DatoDataBase(TAccesos, CampoDB(TAccesos, acc_tipo), "Hoja", CampoDB(TAccesos, acc_contrasenia))
        If Not IsEmpty(passDb) And Not IsNull(passDb) Then
            CachePassHoja = Desencriptar(CStr(passDb))
        End If
    End If
    GetPassHoja = CachePassHoja
End Function


Function ValidUser(user As String, password As String) As Boolean
    ' Bloqueo de StockTech en caso de errores
    On Error GoTo ErrorHandler
    
    Dim realPassword As Variant ' Variant por si la DB devuelve Empty/Null
    Dim beforeCaption As String
    
    ' Por cualquier cosa la validación del usuario permanece en FALSE
    ValidUser = False
    
    ' Evita la comparación de espacios vacíos
    If Trim(user) = "" Or Trim(password) = "" Then
        MsgBox "El usuario y la contraseña son obligatorios.", vbExclamation, "Validación"
        Exit Function
    End If
    
    ' Se obtiene contraseña de la DB
    realPassword = DataBaseUtils.DatoDataBase(TUsuarios, CampoDB(TUsuarios, us_usuario), user, CampoDB(TUsuarios, us_contrasenia))
    
    ' Validamos la extracción de la contraseña
    If IsEmpty(realPassword) Or IsNull(realPassword) Or realPassword = "" Then
        MsgBox "El usuario o la contraseña son incorrectos." & vbNewLine & "Vuelva a intentarlo.", vbExclamation, "Acceso Denegado"
        Exit Function
    End If
    
    If password = Desencriptar(CStr(realPassword)) Then
        
        ' Revisión de la contraseña genérica de primer ingreso
        If CStr(realPassword) = "AFB0CDBAC1B3C0BA3EA2" Then
            MsgBox "Por favor, actualice su contraseña antes de continuar.", vbExclamation, "Seguridad de Cuenta"
            
            beforeCaption = UserLogIn.Caption
            With UpdatePasswordChard
                .Caption = "Reboot Inicio"
                .UserLabel.Caption = user
                .CurrentPass.Enabled = False
                .CurrentPass.BackColor = &H80000004
                .CurrentPass.Text = password
                .Show
            End With
            
            Unload UpdatePasswordChard
            Unload UserLogIn
            UserLogIn.Caption = beforeCaption
            
            Exit Function
        Else
            ValidUser = True
        End If
        
    Else
        MsgBox "El usuario o la contraseña son incorrectos." & vbNewLine & "Vuelva a intentarlo.", vbExclamation, "Acceso Denegado"
    End If
    
    Exit Function

ErrorHandler:
    
    MsgBox "Error de conexión al validar credenciales: " & Err.Description, vbCritical, "Fallo de Sistema"
    ValidUser = False
End Function



Function LevelUser(user As String) As String
    
    LevelUser = DataBaseUtils.DatoDataBase(TUsuarios, CampoDB(TUsuarios, us_usuario), UCase(user), CampoDB(TUsuarios, us_nivel))
    
    If user <> Empty Then
        If LevelUser = Empty Then
            MsgBox "El usuario" & UCase(user) & " No está registrado en la base de datos"
        End If
    End If
End Function
Function UnlockBook()
    ActiveWorkbook.Unprotect GetPassLibro()
End Function
Function LockBook()
    ActiveWorkbook.Protect password:=GetPassLibro(), Structure:=True, Windows:=False
End Function

Function UnlockSheet(sh As Worksheet)
    sh.Unprotect GetPassHoja()
End Function

Function LockSheet(ByVal sh As Worksheet, Optional ByVal Rango As Range)
    Dim pass As String
    
    ' =======================================================
    ' PASO 1: OBTENCIÓN DE CREDENCIALES
    ' =======================================================
    ' Invocamos el motor centralizado para obtener la clave del sistema.
    pass = GetPassHoja()

    ' =======================================================
    ' PASO 2: APERTURA DEL ESCUDO (Unprotect)
    ' =======================================================
    ' Desprotegemos la hoja temporalmente para que VBA tenga permisos
    ' de alterar la propiedad .Locked de las celdas.
    sh.Unprotect pass
    
    ' =======================================================
    ' PASO 3: GESTIÓN DE BLOQUEOS DE CELDA
    ' =======================================================
    ' Al no haber igualado a Nothing en los parámetros, VBA lo evalúa
    ' naturalmente aquí.
    If Rango Is Nothing Then
        ' CASO A: Comportamiento por defecto.
        ' Se sella la hoja entera asegurando que ninguna celda sea editable.
        sh.Cells.Locked = True
    Else
        ' CASO B: Excepciones de escritura.
        ' Primero se sella todo el plano, y luego se perfora el bloqueo
        ' exclusivamente en el rango proporcionado.
        sh.Cells.Locked = True
        Rango.Locked = False
    End If

    ' =======================================================
    ' PASO 4: APLICACIÓN DE PROTECCIÓN INTELIGENTE
    ' =======================================================
    ' UserInterfaceOnly:=True es la directiva más crítica de StockTech:
    ' Bloquea al usuario humano, pero permite que las macros operen sin
    ' tener que desproteger la hoja constantemente.
    If sh.Name = TAuditTrail Then
        sh.Protect password:=pass, _
                   UserInterfaceOnly:=True, _
                   AllowFiltering:=True, _
                   AllowUsingPivotTables:=True, _
                   AllowSorting:=True
    Else
        sh.Protect password:=pass, _
                   UserInterfaceOnly:=True, _
                   AllowFiltering:=True ' Recomendado para vistas estándar como SOLPEDs
    End If
End Function

Function VerificarPrivilegio(user As String, privilegio As String) As Boolean
    Dim dato As Variant
    dato = DataBaseUtils.QueryTableConsult(TUsuarios, TPrivilegios, CampoDB(TUsuarios, us_nivel), privilegio, , CampoDB(TUsuarios, us_usuario), user)
    On Error GoTo SinCoincidencia
    VerificarPrivilegio = dato(6, 0)
    Exit Function
SinCoincidencia:
    MsgBox "Error al buscar privilegio", vbCritical
    Debug.Print ("Debes añadir el privilegio " & privilegio & " a la Base de Datos")
    VerificarPrivilegio = False
    Exit Function
End Function
Sub BloquearSistema(Optional automatico As Boolean = True)
    Dim respuesta As VbMsgBoxResult
    
    ' ¡APAGAMOS EL RELOJ!
    Call DetenerTemporizador
    
    ' 1. Si el bloqueo es manual, confirmamos la acción
    If automatico = False Then
        respuesta = MsgBox("Se bloqueará la sesión de " & UCase(GetCurrentUser) & vbNewLine & _
                           "¿Desea Continuar?", _
                           vbYesNo + vbQuestion, "Bloquear Sistema")
                           
        If respuesta = vbNo Then Exit Sub ' Cancelación de cierre se seción
    End If
    
    ' Cerrando StockTech
    ActiveStocktech = False
    
    GlobalLOG = True
    
    Call App.ResetLabels
    Call DataBaseUtils.CloseDBConnection ' Borrando conexión con Audit
    
    ' Actualización del estado de las etiquetas
    
    ' Registo de AuditTrial
    If automatico = True Then
        CURRENTUSER = ""
        GlobalWHO = "Sistema"
        Call DataBaseUtils.SystemLogDB("Bloqueo automático", "Sesión StockTech")
        MsgBox "StockTech Bloqueado automáticamente por seguridad", vbInformation, "Seguridad"
    Else
        GlobalWHO = GetCurrentUser()
        Call DataBaseUtils.SystemLogDB("Cierre de sesión", "StockTech")
        
        MsgBox "Sesión bloqueada correctamente." & vbNewLine & _
               "Use el botón 'Iniciar Sesión' para reconectar.", vbInformation, "StockTech"
    End If
    ClearCurrentUser
    App.CargarPrivilegiosStockTech (GetCurrentUser)
    On Error Resume Next
    StockTechLabels.Invalidate
    On Error GoTo 0
    
    Set Seguridad.LibroStockTech = Nothing
    
End Sub
Sub IniciarTemporizador()
'    Call DetenerTemporizador
'    HoraCierre = Now + TimeValue("00:01:00")
'    Application.OnTime HoraCierre, "BloquearSistema"
End Sub

Sub DetenerTemporizador()
'    On Error Resume Next
'    Application.OnTime HoraCierre, "BloquearSistema", , False
'    On Error GoTo 0
End Sub

Public Sub SetCurrentUser(ByVal nombreUsuario As String)
    ' Guardamos el nombre en el perfil local de Windows de esta PC específica
    SaveSetting "StockTechApp", "SesionActiva", "UsuarioLogueado", nombreUsuario
End Sub

' Función para LEER al usuario activo (Úsala en tu AuditTrail, Reportes, etc.)
Public Function GetCurrentUser() As String
    ' Leemos el Registro. Si no hay sesión, devuelve ""
    GetCurrentUser = GetSetting("StockTechApp", "SesionActiva", "UsuarioLogueado", "")
End Function

' Función para BORRAR al usuario (Úsala en tu subrutina BloquearSistema)
Public Sub ClearCurrentUser()
    On Error Resume Next
    ' Destruimos la llave del registro al cerrar sesión
    DeleteSetting "StockTechApp", "SesionActiva", "UsuarioLogueado"
    On Error GoTo 0
End Sub
Public Sub MarkStockTechSheet(ByVal ws As Worksheet)
    Dim prop As CustomProperty
    
    On Error Resume Next
    
    ' 1. Limpiamos cualquier marca previa para evitar duplicados en memoria
    For Each prop In ws.CustomProperties
        If prop.Name = "Origin" Then prop.Delete
    Next prop
    
    ' 2. Estampamos la marca única
    ws.CustomProperties.Add Name:="Origin", Value:="StockTech"
    
    On Error GoTo 0
End Sub
Public Sub UnmarkStockTechSheet(ByVal ws As Worksheet)
    Dim prop As CustomProperty
    
    
    On Error Resume Next
    
    For Each prop In ws.CustomProperties
        
        If prop.Name = "Origin" Then
            
            prop.Delete
            
        End If
    Next prop
    
    On Error GoTo 0
End Sub
Public Function GetStockTechSheetName(ByVal ws As Worksheet) As String
    Dim prop As CustomProperty
    
    ' 1. Por defecto, devolvemos una cadena vacía (significa que no tiene la marca)
    GetStockTechSheetName = ""
    
    ' 2. Buscamos en las propiedades personalizadas
    On Error Resume Next
    For Each prop In ws.CustomProperties
        ' Usamos UCase para asegurar la lectura sin importar cómo se guardó
        If UCase(prop.Name) = "ORIGIN" And UCase(prop.Value) = "STOCKTECH" Then
            ' 3. Si la encontramos, asignamos el nombre de la hoja a la función
            GetStockTechSheetName = ws.Name
            Exit For
        End If
    Next prop
    On Error GoTo 0
End Function
Public Function EsDesarrollador() As Boolean
    Dim credencialActual As String
    
    ' Extraemos la credencial de Windows del operador actual
    credencialActual = LCase(Environ("USERNAME"))
    
    ' Comparamos contra la matriz de desarrolladores autorizados
    If credencialActual = "hcruiz" Then
        EsDesarrollador = True
    Else
        EsDesarrollador = False
    End If
End Function
