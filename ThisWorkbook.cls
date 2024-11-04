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
Public clave As String

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Cancel = False
    ThisWorkbook.Save
End Sub

Private Sub Workbook_Open()
    clave = "Rerda2024"
    Dim hojita As Worksheet
    
    ' Se para en la planilla principal
    Sheets(1).Activate
    Range("A1").Activate
    
    ThisWorkbook.Unprotect Password:=clave
    
    'Mostrando y habilitando algunas
    For Each hojita In ThisWorkbook.Worksheets
        If hojita.Visible = False Then
            hojita.Visible = True
        End If
    Next hojita

    ' Ocultando algunas
    Sheets("Datos").Visible = False
    Sheets("LINCE").Visible = False
    
    ' Protegiendo las hojas
    For Each hojita In Application.Worksheets
        hojita.Protect Password:=clave
    Next hojita
    
    'Protegiendo el libro
    ThisWorkbook.Protect Password:=clave

End Sub
