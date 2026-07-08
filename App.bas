Attribute VB_Name = "App"
Option Private Module
Const MODULE_NAME As String = "App"
' Herramienta del núcleo de Windows para copiar direcciones de memoria RAM
#If VBA7 Then
    ' Para Office de 64 bits (Moderno)
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#End If

Global StockTechLabels As IRibbonUI ' Objeto VBA (referencia para cinta de opciones de StockTech)
' Declaramos la variable global para mantener vivo el escuchador
Public StockTechEvents As CEventsApp
Private dictPrivilegios As Object
Private EnProcesoDeRegistro As Boolean 'Requisición para evitar bucles en registro de errores


' Llamar justo después de un Login exitoso
Public Sub StartReadEvents()
    Set StockTechEvents = New CEventsApp
    Set StockTechEvents.StockTechApp = Application
End Sub

' Llamar cuando el sistema se bloquee para ahorrar memoria
Public Sub EndReadEvents()
    Set StockTechEvents = Nothing
End Sub


'Metodo para cargar la cinta de opciones una vez abierto un Excel con la extensión habilitada
Sub StockTechRibbonX(LabelsXML As IRibbonUI)
    Set StockTechLabels = LabelsXML
    
    Dim MemoryPath As String
    MemoryPath = CStr(ObjPtr(LabelsXML))
    
    ThisWorkbook.Names.Add name:="StockTechLabelsID", RefersTo:=MemoryPath, Visible:=False
End Sub
'Metodo para reestablecer la cinta de opciones en caso de errores inesperados
Sub ResetLabels()
    
    If StockTechLabels Is Nothing Then
        
        Dim path As LongPtr
        On Error Resume Next
        
        path = CLngPtr(Replace(ThisWorkbook.Names("StockTechLabelsID").Value, "=", ""))
        On Error GoTo 0
        
        If path <> 0 Then
            Dim clonLabels As Object
            
            CopyMemory clonLabels, path, LenB(path)
            
            Set StockTechLabels = clonLabels
            Set clonLabels = Nothing
        End If
    End If
End Sub
' Metodo de verificacion de identidades autorizadas en el uso de VenTech
Sub GetEnabled(control As IRibbonControl, ByRef returnedVal)
    returnedVal = ActiveStocktech
End Sub
' Método de inicio y apagado de VenTech (Ribbon Callback)
Sub Start(control As IRibbonControl)
    
    If Not ActiveStocktech Then
        ' ==========================================
        ' RUTA DE ENCENDIDO (LOGIN)
        ' ==========================================
        
        ' 1. Validamos que exista un Excel abierto antes de iniciar
        If ActiveWorkbook Is Nothing Then
            Dim respuestaEnv As VbMsgBoxResult
            respuestaEnv = MsgBox("No hay ningún libro abierto para trabajar con StockTech." & vbNewLine & _
                                  "¿Desea crear un libro nuevo en blanco para iniciar la sesión?", _
                                  vbYesNo + vbQuestion, "Entorno no detectado")
            
            If respuestaEnv = vbYes Then
                ' StockTech crea un libro nuevo instantáneamente
                Application.Visible = True
                Workbooks.Add
            Else
                ' Si el usuario dice que no, entonces sí abortamos el inicio
                Exit Sub
            End If
        End If
        
        '==============================
        ' Verificación de reinicio
        '==============================
        Dim user As String
        user = UCase(GetCurrentUser)
        If user <> Empty Then
            Seguridad.BloquearSistema (False)
            ActiveStocktech = True
            Call ResetLabels
            StockTechLabels.Invalidate
        End If
        ' 2. Mostramos tu formulario de carga/login
        StockTech.Show
        
        ' 3. Si el login fue exitoso (tu formulario se encargó de poner GlobalLOG = True)
        If GlobalLOG Then
            ActiveStocktech = True
            
            ' Se genera libro STOCKTECH
            Set Seguridad.LibroStockTech = ActiveWorkbook
            
            ' Se inicia el Patrón OBESERVER
            Call StartReadEvents
            Call Seguridad.IniciarTemporizador
            ' --------------------------------------
            Call Seguridad.LockBook
            MsgBox "Auditoría de StockTech en curso." & vbNewLine & "El sistema está monitoreando la actividad.", vbInformation, "Sesión Iniciada"
        End If
        
    Else
        ' ==========================================
        ' RUTA DE APAGADO (CERRAR SESIÓN MANUAL)
        ' ==========================================
        
        Call Seguridad.BloquearSistema(automatico:=False)
        'Bloquear sistema se encarga de bloquear los labels
        
        Exit Sub
    End If
    
    ' Refrescamos la interfaz gráfica del Ribbon
    ' (El código solo llegará aquí si fue un encendido exitoso o si el usuario canceló el Login a la mitad)
    Call ResetLabels
    StockTechLabels.Invalidate
    
End Sub
Sub CargarPrivilegiosStockTech(user As String)
    Dim nivel As String
    Dim datos As Variant
    Dim i As Long
    
    Set dictPrivilegios = CreateObject("Scripting.Dictionary")
    
    ' Se obtiene Nivel del usuario logeado
    nivel = Seguridad.LevelUser(user)
    
    ' Se obtiene fila de privilegios
    datos = DataBaseUtils.TableDataBase(TPrivilegios, "*", True, CampoDB(TPrivilegios, priv_nivel) & " = " & Chr(34) & nivel & Chr(34))
    
    ' Se crea un diccionario
    If Not IsEmpty(datos) Then
        For i = 2 To UBound(datos)
            ' datos(i,0) es el ID del control, datos(i,1) es el valor booleano
            dictPrivilegios.Add CStr(datos(i, 0)), CBool(datos(i, 1))
        Next i
    End If
End Sub
Sub PrivilegesStockTech(control As IRibbonControl, ByRef returnedVal)
    ' Por defecto, si no está en la base de datos o el diccionario no existe, es False
    returnedVal = False
    
    If Not dictPrivilegios Is Nothing Then
        If dictPrivilegios.Exists(control.id) Then
            returnedVal = dictPrivilegios(control.id)
        End If
    End If
End Sub

Sub OnAction_Ribbon(control As IRibbonControl)
    
    ' Reset de temporizador
    Call Seguridad.IniciarTemporizador
    
    ' Redirigimos al módulo correspondiente según el ID del botón que pulsaron
    Select Case control.id
    
        Case "NewCodsBttn"
            With TipoSolicitud
                .Caption = "Códigos Nvos"
                .Servicio.Visible = False
                .Servicio.Enabled = False
                .Show
            End With
            
        Case "NewSolicitudBttn"
            TipoSolicitud.Show
            
        Case "SOLPEDsbttn"
            Call AccionesRibbon.RegistrarSOLPED
            
        Case "AprobarSOLPED"
            Call AccionesRibbon.AprobarSOLPED
        
        Case "ConsultDBBttn"
            Consulta.Show
            
        Case "PresupuestosBttn"
            AccionesRibbon.CalculosPresupuestos
            
        Case "SolicitudBttn"
            AccionesRibbon.MostrarSolicitudes
            
        Case "SeguimientoBttn"
            SeguimientoSOLPED.Show
            
        Case "AnalisisUserBttn"
            UserDetails.Show
        
        Case "VerTablaBttn"
            VerTabla.Show
            
        Case "UpdateBttn"
            AccionesRibbon.UpdateSheet
            
        Case "EliminarBttn"
            AccionesRibbon.EliminarRegistro
        Case "EditarBttn"
            AccionesRibbon.EditarRegistro
        Case "usersBttn"
            AdminUsers.Show
            
        Case "AuditBttn"
            DataBaseUtils.GetExcelTable TAuditTrail, refresh:=True
            Call Seguridad.LockSheet(ActiveSheet)
    End Select
    
End Sub
Sub Terminar(control As IRibbonControl)
    Seguridad.BloquearSistema (False)
    Call ResetLabels
    StockTechLabels.Invalidate
End Sub

Public Sub SystemError(ByVal detalleError As String)
    
    MsgBox "El sistema StockTech ha encontrado un inconveniente inesperado." & vbCrLf & vbCrLf & _
           "Detalle técnico:" & vbCrLf & detalleError & vbCrLf & vbCrLf & _
           "Este evento ha sido registrado en una bitácora de seguridad para el Administrador del Sistema.", _
           vbCritical, "StockTech - Excepción del Sistema"
    
    Dim rutaError As String
    rutaError = "Error ocurrido en [" & MODULE_NAME & "." & PROC_NAME & "]..."
    detalleError = rutaError & detalleError
    Call RegistrarFallo(detalleError)
End Sub

Private Sub RegistrarFallo(ByVal detalle As String)
    ' Prevención de Bucle Infinito
    If EnProcesoDeRegistro Then Exit Sub
    EnProcesoDeRegistro = True

    ' --- PASO A: LA CAJA NEGRA (TXT Local y Seguro) ---
    Dim rutaArchivo As String
    Dim numArchivo As Integer
    Dim lineaLog As String

    rutaArchivo = ThisWorkbook.path & "\StockTech_MasterLog.txt"
    numArchivo = FreeFile

    lineaLog = "[" & Format(Now(), "yyyy-mm-dd hh:nn:ss") & "] " & _
               "USER: " & UCase(Environ("USERNAME")) & " | " & _
               "ERR: " & Replace(detalle, vbCrLf, " | ")

    On Error Resume Next
    Open rutaArchivo For Append As #numArchivo
    Print #numArchivo, lineaLog
    Close #numArchivo
    On Error GoTo 0

    ' --- PASO B: INTENTO DE AUDIT TRAIL EN DB (Complementario) ---
    Dim tabla As String
    Dim campos As Variant
    Dim valores As Variant
    
    tabla = TSystemTrack
    campos = Array("Fecha", "Usuario", "Lugar", "Error")
    
    valores = Array(Now(), UCase(Environ("USERNAME")), "Excepción Global", detalle)

    If DataBaseUtils.AddRegister(tabla, campos, valores) = False Then
        lineaLog = "[" & Format(Now(), "yyyy-mm-dd hh:nn:ss") & "] " & _
               "USER: " & UCase(Environ("USERNAME")) & " | " & _
               "ERR: Posible error de CONEXIÓN"
    
        On Error Resume Next
        Open rutaArchivo For Append As #numArchivo
        Print #numArchivo, lineaLog
        Close #numArchivo
        On Error GoTo 0
    End If

    EnProcesoDeRegistro = False
End Sub

Public Sub EnfocarStockTechLabels()
    On Error GoTo ErrorHandler
    
    ' 1. INTENTO DE RESURRECCIÓN: Si está vacío, intentamos revivirlo con tu API de Windows
    If StockTechLabels Is Nothing Then
        Call ResetLabels
    End If
    
    ' 2. EJECUCIÓN O ABORTO: Evaluamos de nuevo
    If Not StockTechLabels Is Nothing Then
        ' Ejecutamos el enfoque usando el ID exacto definido en el XML
        StockTechLabels.ActivateTab "tabStocktech"
    Else
        ' Solo llegaremos aquí si incluso el API de Windows falló al intentar recuperarlo
        MsgBox "Se ha perdido la conexión visual con la barra de StockTech." & vbCrLf & _
               "Por favor, guarde sus cambios y reinicie el sistema.", vbExclamation, "StockTech System"
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Usamos tu propio motor centralizado de errores para que quede en la bitácora
    App.SystemError "No se pudo enfocar la pestaña tabStocktech. Error: " & Err.Description
End Sub

Sub ExportarTodoElProyecto()
    Dim oVBProject As Object
    Dim oComponent As Object
    Dim sRuta As String
    Dim iFile As Integer
    Dim sLinea As String
    Dim sContenido As String
    
    ' Carpeta de destino (cámbiala si quieres)
    sRuta = "C:\Users\hcruiz\Desktop\StockTechGit\"
    
    ' Crear carpeta si no existe
    If Dir(sRuta, vbDirectory) = "" Then
        MkDir sRuta
    End If
    
    ' Recorrer todos los componentes del proyecto
    For Each oComponent In ThisWorkbook.VBProject.VBComponents
        Dim sArchivo As String
        
        ' Elegir extensión según tipo
        Select Case oComponent.Type
            Case 1: sArchivo = sRuta & oComponent.name & ".bas"   ' Módulo estándar
            Case 2: sArchivo = sRuta & oComponent.name & ".cls"   ' Clase
            Case 3: sArchivo = sRuta & oComponent.name & ".frm"   ' UserForm
            Case 100: sArchivo = sRuta & oComponent.name & ".cls" ' ThisWorkbook / hojas
        End Select
        
        ' Exportar el componente
        oComponent.Export sArchivo
    Next oComponent
    
    MsgBox "Exportación completa en:" & vbCrLf & sRuta, vbInformation
End Sub
