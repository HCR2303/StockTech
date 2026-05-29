VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CostoRefaccion 
   Caption         =   "Costo de Refacción"
   ClientHeight    =   3384
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "CostoRefaccion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "CostoRefaccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public registroSOLPED As Boolean
Const MODULE_NAME As String = "CostoRefaccion"
Private Sub Costo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' =======================================================
    ' INTERCEPTOR DE TECLADO PARA CAMPOS NUMÉRICOS (DECIMALES)
    ' =======================================================
    
    Select Case KeyAscii
        ' 1. Teclas Permitidas: Números del 0 al 9 (ASCII 48-57) y Retroceso (8)
        Case 48 To 57, 8
            ' Dejamos pasar la tecla con normalidad
            
        ' 2. El Separador Decimal: Punto (46) o Coma (44)
        Case 46, 44
            ' Convertimos automáticamente la coma en punto para estandarizar StockTech
            KeyAscii = 46
            
            ' REGLA DE NEGOCIO: Solo puede existir un punto en toda la cifra
            If InStr(1, Me.Costo.Value, ".") > 0 Then
                ' Si ya hay un punto escrito, aniquilamos este segundo intento
                KeyAscii = 0
            End If
            
        ' 3. Cualquier otra tecla (Letras, símbolos, espacios)
        Case Else
            ' Asignar 0 a KeyAscii destruye la pulsación. El usuario no verá nada en pantalla.
            KeyAscii = 0
    End Select
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

Private Sub RegisCosto_Click()
    Dim vacio As Long: vacio = 0
    Dim cont As control
    For Each cont In Me.Controls
        If TypeName(cont) = "TextBox" Then
            If cont = Empty And cont.Enabled = True Then
                vacio = vacio + 1
            End If
        End If
    Next
    
    If vacio > 0 Then
        MsgBox "Todos los campos son obligatorios", vbExclamation, "Registro de Factura"
        Exit Sub
    End If
    Dim spd As String
    Dim ref As String
    
    spd = Split(Me.SOLPEDLabel.Caption, "SOLPED: ")(1)
    ref = Split(Me.RefaccionLabel.Caption, "Refacción: ")(1)
    
    'Registro en Tabla de SOLPED's
    
    Dim camposU As Variant
    Dim valoresU As Variant
    registroSOLPED = False
    If LCase(ref) Like "serv" & "*" Then
        camposU = Array(CampoDB(TSOLPEDs, spd_orden_compra), CampoDB(TSOLPEDs, spd_no_factura), CampoDB(TSOLPEDs, spd_unidades), CampoDB(TSOLPEDs, spd_costo))
        valoresU = Array(Me.OCompra.Text, Me.NoFactura.Text, 1, CDbl(Me.Costo.Text))
    Else
        camposU = Array(CampoDB(TSOLPEDs, spd_orden_compra), CampoDB(TSOLPEDs, spd_no_factura), CampoDB(TSOLPEDs, spd_unidades), CampoDB(TSOLPEDs, spd_costo))
        valoresU = Array(Me.OCompra.Text, Me.NoFactura.Text, CInt(Me.Unidades.Text), CDbl(Me.Costo.Text))
    End If
    Dim camposWh As Variant
    Dim valoresWh As Variant
    camposWh = Array(CampoDB(TSOLPEDs, SPD_SOLPED), CampoDB(TSOLPEDs, SPD_REFACCION))
    valoresWh = Array(spd, ref)
    
    
    If DataBaseUtils.SetRegister(TSOLPEDs, camposU, valoresU, camposWh, valoresWh) Then
        
        'Registro de actualización de STOCK
        If Not LCase(ref) Like "serv" & "*" Then ' La actualización de STOCK solo ocurre en refacciones distintas a servicio
            Dim stockAnterior As Integer
            Dim idRefaccion As Integer
            stockAnterior = DataBaseUtils.DatoDataBase(TRefacciones, CampoDB(TRefacciones, ref_material), ref, CampoDB(TRefacciones, ref_stock))
            idRefaccion = DataBaseUtils.DatoDataBase(TRefacciones, CampoDB(TRefacciones, ref_material), ref, CampoDB(TRefacciones, ref_id))
        
            If DataBaseUtils.SetDataByID(TRefacciones, CampoDB(TRefacciones, ref_stock), stockAnterior + CInt(Me.Unidades.Text), idRefaccion) Then
                registroSOLPED = True
            End If
        Else
            registroSOLPED = True
        End If
    End If
    Unload Me
    
End Sub

Private Sub UserForm_Activate()
    Dim ref As String
    ref = Split(Me.RefaccionLabel.Caption, "Refacción: ")(1)
    If LCase(ref) Like "serv" & "*" Then
        Me.Unidades.Enabled = False
        Me.Unidades.Visible = False
        Me.UnidadesLabel.Visible = False
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then ' 0 = El usuario presionó la "X"
        registroSOLPED = False
    End If
End Sub
