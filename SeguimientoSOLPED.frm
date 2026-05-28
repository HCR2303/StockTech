VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SeguimientoSOLPED 
   Caption         =   "Registro de SOLPED"
   ClientHeight    =   9828.001
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

Function GetMatrix()
    Dim SOLPEDMatrix As Variant
    Dim sqlQuery As String
    
    ' 1. Construimos el SQL dinámico
    sqlQuery = "SELECT DISTINCT [" & CampoDB(TSOLPEDs, SPD_SOLPED) & "] " & _
               "FROM [" & TSOLPEDs & "] " & _
               "WHERE [" & CampoDB(TSOLPEDs, SPD_SOLPED) & "] IS NOT NULL " & _
               "AND [" & CampoDB(TSOLPEDs, spd_costo) & "] IS NULL"
               
    ' 2. Disparamos la extracción
    SOLPEDMatrix = GetFromSQL(sqlQuery)
    
    ' 3. Validación de vacío
    If IsEmpty(SOLPEDMatrix) Then
        MsgBox "No hay SOLPED's pendientes de registro", vbInformation
        r = MsgBox("¿Desea registrar una nueva SOLPED?", vbYesNo, "Sin Registros Pendientes")
        If r = vbYes Then
            Unload Me
            AccionesRibbon.RegistrarSOLPED
        End If
        Exit Function
    End If
    
    ' La propiedad .Column recibe nativamente la matriz (Cols, Filas) de ADO
    Me.SOLPED.Column = SOLPEDMatrix

End Function

Private Sub Cerrada_Click()
    If Me.Cerrada.Value Then
        Me.ObservacionLabel.Visible = False
        Me.observaciones.Visible = False
    End If
End Sub

Private Sub Proceso_Click()
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
        
        If Me.Cerrada.Value Then
            Me.ObservacionLabel.Visible = False
            Me.observaciones.Visible = False
        End If
        
    End If

End Sub

Private Sub RegistroSeguimiento_Click()
    Dim ctr As control
    If Me.Proceso.Value = False And Me.Cerrada.Value = False Then
        MsgBox "Debe seleccionar y llenar los campos de Estado de SOLPED", vbExclamation, "Estado de SOLPED's"
        Exit Sub
    End If
    
    Dim Refacciones As Variant
        
    If Me.SOLPED.Text = Empty Then
        MsgBox "Debe seleccionar una SOLPED", vbExclamation
        Exit Sub
    End If
        
    
    
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
    If Me.Proceso.Value Then
        Call SOLPEDProceso(Refacciones)
        Unload Me
        Exit Sub
    End If
    If Me.Cerrada.Value Then
        Call SOLPEDCierre(Refacciones)
        Unload Me
        Exit Sub
    End If
    Unload Me
End Sub
Private Sub SOLPEDProceso(Refacciones As Variant)
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
Private Sub SOLPEDCierre(ByRef Refacciones As Variant)
    Me.Hide
    
    Dim spd As String
    Dim ref As String
    Dim i As Long
    spd = Me.SOLPED.Text
    For i = LBound(Refacciones) To UBound(Refacciones)
        ref = Refacciones(i)
        If Not StringUtils.CerrarRefaccion(ref, spd) Then
            MsgBox "Ha ocurrido un error de registro. Intente nuevamente", vbCritical, "Error de ESTADO de SOLPED"
            r = MsgBox("¿Desea volver a intentar?", vbYesNo, "Intentar de Nuevo")
            If r = vbYes Then
                i = i - 1
            Else
                Exit Sub
            End If
        End If
    Next
    GetMatrix
    
    Resultado = Application.Match(spd, Me.SOLPED.List, 0)
        
    If IsError(Resultado) Then
        tabla = TSOLPEDsTrack
    
        campos = Array(CampoDB(tabla, spt_usuario), CampoDB(tabla, spt_fecha), CampoDB(tabla, spt_estado), CampoDB(tabla, spt_solped), CampoDB(tabla, spt_comentario))
        valores = Array(GetCurrentUser, Now(), "CERRADA", Me.SOLPED.Text, Me.observaciones.Text)
        
        If DataBaseUtils.LogDB("Cierre de SOLPED", "SOLPED: " & spd, "Se registra estado de CERRADA de SOLPED: " & spd) Then
            If DataBaseUtils.AddRegister(tabla, campos, valores) Then
                MsgBox "La SOLPED: " & spd & " ya no tiene refacciones por registrar y se declarará CERRADA ", vbInformation, "SOLPED en Proceso"
                Unload Me
            Else
                MsgBox "Error en registro de SOLPED CERRADA. Contacte a Administrador", vbCritical
            End If
        End If
    End If
        
End Sub

Private Sub SOLPED_Click()
    Me.Proceso.Value = False
    Me.Cerrada.Value = False
    
    Me.ObservacionLabel.Visible = False
    Me.observaciones.Visible = False
    
    fila = SOLPED.ListIndex
    
    If fila <> -1 Or SOLPED <> Empty Then
        CargarDatosSOLPED (SOLPED.Value)
    End If
End Sub

Private Sub UserForm_Activate()
    GetMatrix
End Sub
Public Sub CargarDatosSOLPED(ByVal idSolpedSeleccionada As String)
    Dim arrSpd As Variant
    Dim arrEqu As Variant
    Dim arrCta As Variant
    Dim arrProv As Variant
    
    ' 1. Traemos SOLO la fila de esa SOLPED (Carga Selectiva)
    
    arrSpd = DataBaseUtils.TableDataBase(TablasDB.TSOLPEDs, "*", False, CampoDB(TSOLPEDs, SPD_SOLPED) & " = '" & idSolpedSeleccionada & "' AND " & _
                                        CampoDB(TSOLPEDs, spd_costo) & " IS NULL")
    
    If IsEmpty(arrSpd) Then Exit Sub
    
    ' 2. Llenamos los TextBox y Captions de forma segura (Sin Resume Next)
    Dim valEquipo As Variant
    Dim valProveedor As Variant
    Dim valUnico As Variant
    Dim valCuenta As Variant
    
    ' Extraemos las matrices de valores únicos
    valEquipo = DataBaseUtils.UniqueValues(CNameSOLPEDs.spd_codigo_equipo, arrSpd)
    valProveedor = DataBaseUtils.UniqueValues(CNameSOLPEDs.spd_proveedor, arrSpd)
    valUnico = DataBaseUtils.UniqueValues(CNameSOLPEDs.spd_proveedor_unico, arrSpd)
    valCuenta = DataBaseUtils.UniqueValues(CNameSOLPEDs.spd_cuenta, arrSpd)
    
    ' Asignamos el índice 0 a la interfaz (solo si la matriz no está vacía)
    If Not IsEmpty(valEquipo) Then Me.CodigoEquipo.Caption = valEquipo(0)
    If Not IsEmpty(valProveedor) Then Me.ProveedorName.Caption = valProveedor(0)
    If Not IsEmpty(valCuenta) Then
        Me.NumeroCuenta.Caption = valCuenta(0)
        arrCta = CamposTablaDB(TCuentas)
        Me.CountName.Caption = DataBaseUtils.DatoDataBase(TCuentas, arrCta(cta_id_cuenta), valCuenta(0), arrCta(cta_nombre))
    End If
    
    ' Para propiedades Booleanas (True/False) como Visible, hacemos una conversión estricta
    If Not IsEmpty(valUnico) Then
        Me.Unico.Visible = CBool(valUnico(0))
    End If
    ' Relleno para frame de EQUIPOS
    arrEqu = DataBaseUtils.TableDataBase(TEquipos, "*", False, "Código = '" & Me.CodigoEquipo.Caption & "'")
    If Not IsEmpty(arrEqu) Then
        Me.MarcaEquipo.Caption = arrEqu(CNameEquipos.equ_marca, 0)
        Me.ModeloEquipo.Caption = arrEqu(CNameEquipos.equ_modelo, 0)
        Me.EquipName.Caption = arrEqu(CNameEquipos.equ_nombre, 0)
        Dim Criticidad As String
        Criticidad = LCase(arrEqu(CNameEquipos.equ_criticidad, 0)) & ""
        Select Case Criticidad
            Case "alto"
                CriticidadRed.Visible = True
                CriticidadYellow.Visible = False
                CriticidadGreen.Visible = False
            Case "medio"
                CriticidadRed.Visible = False
                CriticidadYellow.Visible = True
                CriticidadGreen.Visible = False
            Case Else
                CriticidadRed.Visible = False
                CriticidadYellow.Visible = False
                CriticidadGreen.Visible = True
        End Select
        Dim camposUbic As Variant
        camposUbic = TablasDB.CamposTablaDB(TUbicaciones)
        Dim ubic As String
        ubic = DataBaseUtils.DatoDataBase(TUbicaciones, camposUbic(ubi_id_ubicacion), arrEqu(equ_id_ubicacion, 0), camposUbic(ubi_ubicacion))
        Me.Ubicacion.Caption = ubic
    End If
    
    ' Relleno para frame de PROVEEDORES
    Dim codigoProv As Variant
    Dim camposProv As Variant

    camposProv = TablasDB.CamposTablaDB(TProveedores)
    codigoProv = DataBaseUtils.DatoDataBase(TProveedores, camposProv(prv_nombre), Me.ProveedorName.Caption, camposProv(prv_codigo))
    If Not IsEmpty(codigoProv) Then
        Me.CodigoProveedor.Caption = CStr(codigoProv)
    End If
    
    ' =========================================================
    ' Relleno para frame de REFACCIONES
    ' =========================================================
    Dim valRefacciones As Variant
    
    ' Usamos el procesamiento en memoria RAM (Ruta B de tu UniqueValues)
    valRefacciones = DataBaseUtils.UniqueValues(CNameSOLPEDs.SPD_REFACCION, arrSpd)
    
    Me.RefaccionesCods.Clear
    
    If Not IsEmpty(valRefacciones) Then
        ' Prevención del Error 9 (Subíndice fuera del intervalo)
        If UBound(valRefacciones) = 0 Then
            ' Si la SOLPED tiene exactamente 1 sola refacción
            Me.RefaccionesCods.AddItem valRefacciones(0)
        Else
            ' Si la SOLPED tiene múltiples refacciones distintas
            Me.RefaccionesCods.List = valRefacciones
        End If
        
    End If
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
