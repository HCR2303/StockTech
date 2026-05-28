VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calendario 
   Caption         =   "Selecciona la Fecha"
   ClientHeight    =   5700
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3600
   OleObjectBlob   =   "Calendario.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private matrizBotones(1 To 42) As New CCalendarButton
Const MODULE_NAME As String = "Calendario"
Public fechaSeleccionada As Date
Private Sub GetDias(ByVal Mes As Integer, ByVal anio As Integer)
    Dim primerDiaMes As Date
    Dim diaSemanaInicio As Integer
    Dim diasMesActual As Integer
    Dim diasMesAnterior As Integer
    Dim i As Integer
    Dim contadorDias As Integer
    Dim btn As MSForms.control
    
    ' =======================================================
    ' 1. CÁLCULO DE LÍMITES Y POSICIONES
    ' =======================================================
    ' Obtenemos la fecha exacta del día 1 del mes solicitado
    primerDiaMes = DateSerial(anio, Mes, 1)
    
    ' Descubrimos en qué día de la semana cae ese día 1.
    diaSemanaInicio = Weekday(primerDiaMes, vbMonday) - 2
    
    ' Calculamos cuántos días tiene el mes actual y el mes anterior (Truco del Día Cero)
    diasMesActual = Day(DateSerial(anio, Mes + 1, 0))
    diasMesAnterior = Day(DateSerial(anio, Mes, 0))
    
    
    Application.ScreenUpdating = False
    
    For i = 1 To 42
        ' Enlazamos la variable dinámica con el control físico del UserForm
        Set btn = Me.Controls("dia" & i)
        Set matrizBotones(i).btnDia = Me.Controls("dia" & i)
        Set matrizBotones(i).frmPadre = Me
        ' Fase A: Días residuales del MES ANTERIOR
        If i < diaSemanaInicio Then
            btn.Caption = CStr(diasMesAnterior - diaSemanaInicio + i + 1)
            btn.Enabled = False
            btn.BackColor = RGB(240, 240, 240) ' Gris inactivo
            btn.ForeColor = RGB(160, 160, 160)
            
        ' Fase B: Días activos del MES ACTUAL
        ElseIf i >= diaSemanaInicio And i < (diaSemanaInicio + diasMesActual) Then
            contadorDias = i - diaSemanaInicio + 1
            btn.Caption = CStr(contadorDias)
            btn.Enabled = True
            btn.BackColor = RGB(255, 255, 255) ' Blanco activo
            btn.ForeColor = RGB(0, 0, 0)
            
            ' Foco de Interfaz (UX): Resaltar el día de HOY
            If DateSerial(anio, Mes, contadorDias) = Date Then
                btn.BackColor = RGB(28, 64, 94) ' Azul corporativo de StockTech
                btn.ForeColor = RGB(255, 255, 255)
            End If
            
        ' Fase C: Días adelantados del MES SIGUIENTE
        Else
            btn.Caption = CStr(i - (diaSemanaInicio + diasMesActual) + 1)
            btn.Enabled = False
            btn.BackColor = RGB(240, 240, 240) ' Gris inactivo
            btn.ForeColor = RGB(160, 160, 160)
        End If
    Next i
    
    Application.ScreenUpdating = True
End Sub

Private Sub Año_Change()
    If Mes.Value <> Empty And Año.Value <> Empty Then
        Call GetDias(Mes.ListIndex - 1, Año.Value)
    End If
End Sub

Private Sub InputDate_Click()
    fechaSeleccionada = CDate(Me.fecha.Caption)
    Me.Hide
End Sub

Private Sub Mes_Change()
    If Mes.Value <> Empty And Año.Value <> Empty Then
        Call GetDias(Mes.ListIndex - 1, Año.Value)
    End If
End Sub

Private Sub UserForm_Activate()
    Dim i As Integer
    
    ' =======================================================
    ' 1. CARGA DINÁMICA DE MESES
    ' =======================================================
    For i = 1 To 12
        Me.Mes.AddItem StrConv(MonthName(i), vbProperCase)
    Next i
    
    ' =======================================================
    ' 2. CARGA DINÁMICA DE AÑOS (Ventana de Tiempo)
    ' =======================================================
    
    For i = Year(Date) - 10 To Year(Date) + 10
        Me.Año.AddItem i
    Next i
    
    ' Selección de fecha en curso
    Me.Mes.ListIndex = Month(Date) - 1
    Me.Año.Value = Year(Date)
    Me.fecha.Caption = Date
End Sub
Public Sub clicDia(ByVal seleccion As String)
    Dim anio As Integer
    Dim dia As Integer
    Dim Mes As Integer
    anio = Calendario.Año.Value
    Mes = Calendario.Mes.ListIndex + 1
    dia = CInt(seleccion)
    Calendario.fecha.Caption = DateSerial(anio, Mes, dia)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = vbFormControlMenu Then ' 0 = El usuario presionó la "X"
        fechaSeleccionada = Empty
    End If
End Sub
