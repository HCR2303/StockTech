VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEventsApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const MODULE_NAME As String = "CEventsApp"
Public WithEvents StockTechApp As Application
Attribute StockTechApp.VB_VarHelpID = -1

Private Sub StockTechApp_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)
    If ActiveStocktech = True Then
        ' 1. Verificamos que el ancla exista
        If Not Seguridad.LibroStockTech Is Nothing Then
            ' 2. ¿El clic fue en el libro anclado?
            If sh.Parent.name = Seguridad.LibroStockTech.name Then
                Call Seguridad.IniciarTemporizador
            End If
        End If
    End If
End Sub

Private Sub StockTechApp_SheetChange(ByVal sh As Object, ByVal Target As Range)
    If ActiveStocktech = True Then
        If Not Seguridad.LibroStockTech Is Nothing Then
            If sh.Parent.name = Seguridad.LibroStockTech.name Then
                Call Seguridad.IniciarTemporizador
            End If
        End If
    End If
End Sub
