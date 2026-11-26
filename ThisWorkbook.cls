VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    Application.ScreenUpdating = False
    Call Proteger
    Application.ScreenUpdating = True
    Worksheets("DOMICILIO").Activate
    Worksheets("DOMICILIO").Range("C29").Activate
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Guardar automáticamente sin mostrar cuadros de diálogo
    Application.ScreenUpdating = False
    Cancel = False
    If Not Me.Saved Then
        Me.Close SaveChanges:=True
    End If
    Application.ScreenUpdating = True
End Sub
