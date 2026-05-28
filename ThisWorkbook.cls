VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
    Call DataBaseUtils.resetGlobals
    If Seguridad.EsDesarrollador Then Exit Sub
    Call VerificarActualizacionStockTech
End Sub
Private Sub VerificarActualizacionStockTech()
    Dim rutaMaestraRed As String
    Dim fechaVersionLocal As Date
    Dim fechaVersionRed As Date
    Dim respuesta As VbMsgBoxResult
    
    ' Ruta de tu repositorio maestro
    rutaMaestraRed = "\\silmxtolf01\Ingenieria y Mantenimiento\STOCKTECH\StockTech.xlam"
    
    ' Verificamos conexión al servidor silenciosamente
    On Error Resume Next
    If Dir(rutaMaestraRed) = "" Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    fechaVersionLocal = FileDateTime(ThisWorkbook.FullName)
    fechaVersionRed = FileDateTime(rutaMaestraRed)
    
    ' Si hay actualización disponible
    If fechaVersionRed > fechaVersionLocal Then
        
        respuesta = MsgBox("Se ha detectado una nueva versión de StockTech." & vbNewLine & vbNewLine & _
                           "¿Desea actualizar el sistema ahora? (Se requiere cerrar Excel)", _
                           vbYesNo + vbQuestion, "Actualización Disponible")
                           
        If respuesta = vbYes Then
            ' Lanzamos el proceso de autoupdate
            Call EjecutarAutoUpdate(rutaMaestraRed, ThisWorkbook.FullName)
        End If
        
    End If
End Sub
Private Sub EjecutarAutoUpdate(ByVal rutaOrigen As String, ByVal rutaDestino As String)
    Dim rutaBat As String
    Dim numArchivo As Integer
    
    ' 1. Definimos la ruta del archivo temporal en la carpeta segura de Windows (%temp%)
    rutaBat = Environ("Temp") & "\UpdaterStockTech.bat"
    
    ' 2. Fabricamos el script .bat línea por línea
    numArchivo = FreeFile
    Open rutaBat For Output As #numArchivo
    
    ' Ocultamos los comandos de consola
    Print #numArchivo, "@echo off"
    
    ' Bucle industrial de sobrescritura:
    ' Intenta copiar repetidamente hasta que Excel libere por completo el archivo local
    Print #numArchivo, ":ReintentarCopiar"
    Print #numArchivo, "copy /Y """ & rutaOrigen & """ """ & rutaDestino & """"
    Print #numArchivo, "if %errorlevel% neq 0 ("
    Print #numArchivo, "    timeout /t 1 >nul"
    Print #numArchivo, "    goto ReintentarCopiar"
    Print #numArchivo, ")"
    
    ' Reiniciamos Excel automáticamente tras una copia exitosa
    Print #numArchivo, "start excel.exe"
    
    ' Autodestrucción del script temporal para no dejar basura en la PC
    Print #numArchivo, "del ""%~f0"""
    
    Close #numArchivo
    
    ' 3. Ejecutamos el archivo en segundo plano (vbHide hace que la pantalla negra de CMD no se vea)
    Shell rutaBat, vbHide
    
    ' 4. Forzamos el cierre de Excel para liberar el archivo .xlam
    ' Nota operativa: Si el usuario tiene otros archivos de Excel sin guardar,
    ' el sistema nativo le preguntará "Desea guardar los cambios" de forma segura.
    Application.Quit
End Sub
