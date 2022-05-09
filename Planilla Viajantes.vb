Option Explicit
Sub Generar_Rotulo()
' ============================================================
' GENERA UN ROTULO Y PLANILLA PARA IMPRIMIR VENTA DE UN 
' VIAJANTE O SUCURSASL. CONTROLA QUE LAS SUBCARPETAS SE HAYAN
' CREADO. SI NO, LAS CREA. TAMBIÉN GUARDA LOS ARCHIVOS CON
' NOMBRE DE LOS CLIENTES Y FECHA, CADA RESPECTIVA SUBCARPETA.
' ============================================================

	' Protegiendo el archivo de cambios innecesarios
	ThisWorkbook.Protect
	If Worksheets("Lince").Visible = True Then
		Worksheets("Lince").Visible = False
	End If
	
	' Definiendo variables a usar
	Dim ruta As String
	Dim rutaViajante As String
	Dim viajante As String
	Dim cliente As String
	Dim fecha As String
	
	
	' Asigando los valores de acuerdo a la planilla donde esté parado
	ruta = ThisWorkbook.Path & "\Ventas de Viajantes"
	viajante = ActiveSheet.Range("X1").Value
	fecha = Day(Date) & "-" & Month(Date) & "-" & Year(Date)
	
	' Controlando y validando el nombre del cliente
	cliente = ActiveSheet.Range("X2").Value
	
	'Controlando si la carpeta existe, de lo contrario, crearla en local
	If Dir(ruta, vbDirectory) = "" Then
		MkDir (ruta)
	End If
	
	' Controlando la carpeta del viajante
	rutaViajante = ruta & "\" & viajante
	If Dir(rutaViajante, vbDirectory) = "" Then
		MkDir (rutaViajante)
	End If
	
	' Guardando el archivo
	ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
		rutaViajante & "\" & cliente & ". " & fecha & ".pdf", Quality:=xlQualityStandard, _
		IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
		True
End Sub
