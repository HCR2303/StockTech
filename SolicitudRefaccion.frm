VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SolicitudRefaccion 
   Caption         =   "Solicitud de Refacción"
   ClientHeight    =   4392
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5388
   OleObjectBlob   =   "SolicitudRefaccion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "SolicitudRefaccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MODULE_NAME As String = "SolicitudRefaccion"
Private RefaccionMatrix As Variant
Function GetMatrix()
    
    Dim RefaccionList() As Variant
    Dim Filas As Long
    
    RefaccionMatrix = DataBaseUtils.TableDataBase("Refacciones", "*")
    
    If IsEmpty(RefaccionMatrix) Then
        MsgBox "Error de obteción de datos de REFACCIONES", vbCritical
        Exit Function
    End If
    
    Filas = UBound(RefaccionMatrix, 2)
    ReDim RefaccionList(0 To Filas)
    For i = 0 To Filas
        RefaccionList(i) = RefaccionMatrix(2, i) & ""
    Next
    
    Me.RefaccionName.List = RefaccionList
    
End Function


Private Sub RefaccionName_Change()
    fila = RefaccionName.ListIndex
    
    If fila = -1 Or RefaccionName = Empty Then
        Rrojo.Visible = False
        Ramarillo.Visible = False
        Rverde.Visible = False
        Nulo.Visible = False
        Exit Sub
    End If
    
    Me.RefaccionCodigo.Caption = RefaccionMatrix(1, fila)
    Stock.Caption = RefaccionMatrix(4, fila) & " " & RefaccionMatrix(3, fila)
    Criticidad = RefaccionMatrix(6, fila)
    Select Case Criticidad
        Case "alto"
            Rrojo.Visible = True
            Ramarillo.Visible = False
            Rverde.Visible = False
            Nulo.Visible = False
        Case "medio"
            Rrojo.Visible = False
            Ramarillo.Visible = True
            Rverde.Visible = False
            Nulo.Visible = False
        Case "bajo"
            Rrojo.Visible = False
            Ramarillo.Visible = False
            Rverde.Visible = True
            Nulo.Visible = False
        Case Else
            Nulo.Visible = True
        End Select
End Sub

Private Sub RefaccionName_Click()
    Call RefaccionName_Change
End Sub

Private Sub SCodigo_Click()
    If SCodigo.Value Then
        SCodigoLabel.Visible = True
        Exit Sub
        
    End If
    SCodigoLabel.Visible = False
End Sub

Private Sub SolicitarRefaccion_Click()
    Dim tabla As String
    Dim campos As Variant
    Dim valores As Variant
    Dim user As String
    'On Error GoTo ErrorDatos
    
    If Me.SCodigo.Value Then
        If Me.Unidades = Empty Then
             MsgBox "Debe completar el campo de unidades a solicitar", vbExclamation, "Falta de Unidades"
             Exit Sub
        End If
        NewRefaction.Caption = "Refacción Sin Código"
        NewRefaction.Unidades.Value = Me.Unidades.Value
        Unload Me
        NewRefaction.Show
        Exit Sub
    End If
    If RefaccionCodigo.Caption = Empty Or Unidades.Value = Empty Then
        MsgBox "Uno o mas campos del formulario está vacío. " & vbNewLine & "Complete para continuar", vbInformation
        Exit Sub
    End If
    campos = Array(CampoDB(TSolicitudes, sol_fecha), CampoDB(TSolicitudes, sol_usuario), CampoDB(TSolicitudes, sol_codigo), CampoDB(TSolicitudes, sol_tipo), CampoDB(TSolicitudes, sol_unidades), CampoDB(TSolicitudes, sol_realizada), CampoDB(TSolicitudes, sol_mantenimiento))
    valores = Array(Now(), GetCurrentUser, RefaccionCodigo.Caption, "REFACCIÓN", CInt(Unidades.Value), False, StringUtils.SetMantenimiento)
    
    If DataBaseUtils.LogDB("Solicitud", "Refacción: " & RefaccionCodigo.Caption, "Se crea solicitud de refacción: " & RefaccionCodigo.Caption) Then
        If DataBaseUtils.AddRegister(TSolicitudes, campos, valores) Then
            MsgBox "Registro exitoso", vbInformation, "Registro de Refacción"
            Call DataBaseUtils.GetExcelTable(TSolicitudes, refresh:=True)
            Call StringUtils.FiltrarTabla(TSolicitudes, CampoDB(TSolicitudes, sol_realizada), False)
            Call Seguridad.LockSheet(ActiveSheet)
            Unload Me
        End If
    End If
    
    'On Error GoTo 0
    Exit Sub
    
ErrorDatos:
    MsgBox "Error de campos o valores en la tabla: " & tabla, vbCritical
    Exit Sub

End Sub
Private Sub Unidades_keypress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' =======================================================
    ' INTERCEPTOR DE TECLADO PARA CAMPOS NUMÉRICOS (DECIMALES)
    ' =======================================================
    
    Select Case KeyAscii
        ' 1. Teclas Permitidas: Números del 0 al 9 (ASCII 48-57) y Retroceso (8)
        Case 48 To 57, 8
            ' Dejamos pasar la tecla con normalidad
            
        Case Else
            ' Asignar 0 a KeyAscii destruye la pulsación. El usuario no verá nada en pantalla.
            KeyAscii = 0
    End Select
End Sub

Private Sub UserForm_Activate()
    Call GetMatrix
End Sub
