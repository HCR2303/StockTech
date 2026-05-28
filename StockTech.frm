VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StockTech 
   ClientHeight    =   2472
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "StockTech.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "StockTech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MODULE_NAME As String = "StockTech"
Sub Sleep(milisegundos As Single)
    Dim final As Double
    final = Timer + (milisegundos / 1000)
    Do While Timer < final
        DoEvents
    Loop
End Sub

Private Sub UserForm_Activate()
    Dim i As Integer
    Dim conn As ADODB.Connection
    Dim rutaManual As String
    Dim conexionExitosa As Boolean
    
    ' =======================================================
    ' FASE 1: PREPARACIÓN DE INTERFAZ
    ' =======================================================
    Loading.BackColor = RGB(28, 64, 94)
    Loading.ForeColor = RGB(255, 255, 255)
    WorkLabel.Caption = "Iniciando módulos de StockTech..."
    
    ' Primera carga visual (0% al ~33%)
    For i = 40 To 56
        Loading.Width = i
        Loading.Caption = Format((i / 168) * 100, "0") & " %"
        DoEvents
        Sleep 20
    Next
    
    ' =======================================================
    ' FASE 2: CONEXIÓN A LA BASE DE DATOS (BUCLE SEGURO)
    ' =======================================================
    WorkLabel.Caption = "Verificando integridad y ruta de base de datos..."
    DoEvents
    
    conexionExitosa = False
    Do While Not conexionExitosa
        
        ' 1. Verificamos que exista una ruta física válida mediante tu nueva función DB()
        If DataBaseUtils.DB = "" Then
            ' Si entra aquí, falló tanto la ruta local (.xlam) como la del Registro
            MsgBox "No se localizó la base de datos de StockTech en el directorio local ni en el registro." & vbNewLine & _
                   "Por favor, localice el archivo 'StockTech DB.accdb' manualmente.", vbExclamation, "Base de Datos no encontrada"
            
            rutaManual = BuscarArchivoDB()
            
            If rutaManual <> "" Then
                ' Encriptamos y guardamos la ruta de contingencia
                SaveSetting "StockTech", "Conexion", "RutaDB", Seguridad.Encriptar(rutaManual)
                ' Forzamos al bucle a reevaluar la función DB() desde cero
                GoTo SiguienteIntento
            Else
                MsgBox "Operación cancelada. El sistema se cerrará por seguridad.", vbCritical, "Cierre de Sistema"
                Call CierreForzado
                Exit Sub
            End If
        End If
        
        ' 2. Si la ruta es válida, intentamos el anclaje del motor ADO
        WorkLabel.Caption = "Estableciendo conexión con el servidor..."
        DoEvents
        
        On Error Resume Next
        Set conn = DataBaseUtils.GetDBConnection()
        On Error GoTo 0
        
        If Not conn Is Nothing Then
            If conn.State = 1 Then
                conexionExitosa = True
                Exit Do
            End If
        End If
        
        ' 3. Si llega a este punto: la ruta física existe, pero el archivo ADO rebotó la conexión
        MsgBox "Se encontró el archivo de base de datos, pero el motor rechazó la conexión." & vbNewLine & _
               "Verifique que el archivo no se encuentre en uso exclusivo o dañado.", vbCritical, "Fallo de Motor ADO"
        Call DataBaseUtils.CloseDBConnection
        Call CierreForzado
        Exit Sub
        
SiguienteIntento:
    Loop
    
    ' =======================================================
    ' FASE 3: CARGA EXITOSA Y APERTURA
    ' =======================================================
    WorkLabel.Caption = "Acceso concedido. Cargando interfaz..."
    
    ' Unificamos la carga visual restante (del ~34% al 100%)
    For i = 57 To 168
        Loading.Width = i
        Loading.Caption = Format((i / 168) * 100, "0") & " %"
        DoEvents
        Sleep 20
    Next
    
    Sleep 600
    Unload Me ' Destruimos la pantalla de carga
    
    ' =======================================================
    ' FASE 4: AUTENTICACIÓN Y ORQUESTACIÓN
    ' =======================================================
    
    With UserLogIn
        .Caption = "Inicio de Sesión"
        .Show
    End With
    
    ' Si el Login fue exitoso
    If GlobalLOG Then
        If DataBaseUtils.LogDB("Inicio de sesión", "StockTech", "Inicio de sesión de usuario: " & UCase(GetCurrentUser)) Then
            MsgBox "Bienvenido a StockTech" & vbNewLine & UCase(GetCurrentUser), vbInformation, "Bienvenido"
            
            Call Seguridad.IniciarTemporizador
            
            ' CORRECCIÓN DE ARQUITECTURA: Aislar variables del proyecto correcto
            ' Reemplazar ActiveVentech = True por la correspondiente a StockTech
            ActiveStocktech = True
            
            Set Seguridad.LibroStockTech = ActiveWorkbook
            If Not ThisWorkbook.ProtectStructure Then
                Call Seguridad.LockBook
            End If
        End If
        App.CargarPrivilegiosStockTech (GetCurrentUser)
        App.ResetLabels
        StockTechLabels.Invalidate
        App.EnfocarStockTechLabels
    Else
        MsgBox "Acceso no autorizado. El sistema se cerrará.", vbExclamation, "Acceso Denegado"
        Call CierreForzado
    End If
End Sub
Private Function BuscarArchivoDB() As String
    Dim fDialog As FileDialog
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .Title = "Seleccionar Base de Datos StockTech"
        .Filters.Clear
        .Filters.Add "Access Database", "*.accdb; *.mdb"
        If .Show = -1 Then BuscarArchivoDB = .SelectedItems(1)
    End With
End Function

Public Sub CierreForzado()
    Dim frm As Object
    
    ' =========================================================
    ' 1. RUPTURA DE CONEXIONES EXTERNAS
    ' =========================================================
    
    On Error Resume Next
    Call DataBaseUtils.CloseDBConnection
    On Error GoTo 0
    
    ' =========================================================
    ' 2. LIMPIEZA DE CREDENCIALES Y ESTADO GLOBAL
    ' =========================================================
    
    DataBaseUtils.resetGlobals
    
    ' =========================================================
    ' 3. ACTUALIZACIÓN DE LA INTERFAZ (RIBBON)
    ' =========================================================
    
    On Error Resume Next
    App.ResetLabels
    StockTechLabels.Invalidate
    On Error GoTo 0
    
    ' =========================================================
    ' 4. DESTRUCCIÓN DE INTERFACES GRÁFICAS (USERFORMS)
    ' =========================================================
   
    On Error Resume Next
    For Each frm In VBA.UserForms
        Unload frm
    Next frm
    On Error GoTo 0
    
End Sub

