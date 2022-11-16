Option Explicit

' Definiendo unas variables globales a usar
Dim cp As String
Dim codigoNis As String
Dim ruta As String
Dim rutaViajante As String
Dim viajante As String
Dim cliente As String
Dim fecha As String
Sub Generar_Rotulo()
' ============================================================
' GENERA UN ROTULO Y PLANILLA PARA IMPRIMIR VENTA DE UN
' VIAJANTE O SUCURSASL. CONTROLA QUE LAS SUBCARPETAS SE HAYAN
' CREADO. SI NO, LAS CREA. TAMBIÉN GUARDA LOS ARCHIVOS CON
' NOMBRE DE LOS CLIENTES Y FECHA, CADA RESPECTIVA SUBCARPETA.
' ============================================================

' Dando formato necesario
' Formato de impresión
With ActiveSheet.PageSetup
	.Orientation = xlLandscape
	.PaperSize = xlPaperA4
	.LeftMargin = Application.CentimetersToPoints(0.64)
	.RightMargin = Application.CentimetersToPoints(0.64)
	.TopMargin = Application.CentimetersToPoints(2.5)
	.BottomMargin = Application.CentimetersToPoints(1.91)
	.HeaderMargin = Application.CentimetersToPoints(0.76)
	.FooterMargin = Application.CentimetersToPoints(0.76)
	.CenterHorizontally = True
	.CenterVertically = False
	.PrintArea = ActiveSheet.Range("A1:AA36")
	.Zoom = False
	.FitToPagesTall = 1
	.FitToPagesWide = 1
	'.CenterHeader = "&B&20&F"
End With
	
	' Ocultando una hoja innecesaria
	If Worksheets("Lince").Visible = True Then
		Worksheets("Lince").Visible = False
	End If
	If Worksheets("Datos").Visible = True Then
		Worksheets("Datos").Visible = False
	End If
	
	' Protegiendo el resto de las hojas de cambios innecesarios
	Dim cuenta As Integer
	For cuenta = 1 To Worksheets.Count
		Worksheets(cuenta).Protect
	Next
	
	' Pretegiendo el libraco
	ActiveWorkbook.Protect Password:="Rerda", Structure:=True, Windows:=True
	
	' Asigando los valores de acuerdo a la planilla donde esté parado
	ruta = ThisWorkbook.Path & "\Ventas de Viajantes"
	viajante = ActiveSheet.Range("X1").Value
	fecha = Day(Date) & "-" & Month(Date) & "-" & Year(Date)
	cliente = ActiveSheet.Range("X2").Value
	cp = ActiveSheet.Range("X5").Value

	
	'Controlando si la carpeta existe, de lo contrario, crearla en local
	If Dir(ruta, vbDirectory) = "" Then
		MkDir (ruta)
	End If
	
	' Controlando la carpeta del viajante
	rutaViajante = ruta & "\" & viajante
	If Dir(rutaViajante, vbDirectory) = "" Then
		MkDir (rutaViajante)
	End If
	
	
	' Controlando al vendedor o viajante
	If viajante = "" Then
		MsgBox ("¿Y tu nombre como vendedor, viajante o sucursal?")
		ActiveSheet.Range("X1").Select
	
	' Controlando apellido, nombre razón social del cliente
	ElseIf ActiveSheet.Range("X2").Value = "" Then
		MsgBox ("¿Y el Apellido/Nombre o razón social del cliente?")
		ActiveSheet.Range("X2").Select
	
	' Controlando la dirección para el caso de envío a domicilio
	ElseIf ActiveSheet.Name = "A Domicilio" And ActiveSheet.Range("X3").Value = "" Then
		MsgBox ("¿Y la dirección de destino? Calle, Nº, piso, dpto, etc...")
		Sheets("A Domicilio").Range("X3").Select
		
	' Controlando el código postal
	ElseIf ActiveSheet.Range("X5").Value = "" Then
		MsgBox ("¿Y el código postal?")
		ActiveSheet.Range("X5").Select
	
	' Controlando el DNI/CUIT
	ElseIf ActiveSheet.Range("X6").Value = "" Then
		MsgBox ("¿Y el DNI o CUIL/CUIT del cliente?")
		ActiveSheet.Range("X6").Select
	
	' Controlando el teléfono de contacto
	ElseIf ActiveSheet.Range("X7").Value = "" Then
		MsgBox ("¿Y el teléfono o celular del cliente?")
		ActiveSheet.Range("X7").Select
	
	' Controlando la ciudad o pueblo
	ElseIf ActiveSheet.Range("X8").Value = "" Then
		MsgBox ("¿Y la ciudad o pueblo?")
		ActiveSheet.Range("X8").Select
	
	' Controlando la Provincia
	ElseIf ActiveSheet.Range("X9").Value = "" Then
		MsgBox ("¿Y la provincia?")
		ActiveSheet.Range("X9").Select
	
	Else
		' Mostrar dónde se guardó el archivo
		MsgBox ("Se guardó una copia PDF en " & rutaViajante & "\" & cliente & ". " & fecha & ".pdf")
		
		' Guardando el archivo
		ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
			rutaViajante & "\" & cliente & ". " & fecha & ".pdf", Quality:=xlQualityStandard, _
			IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
			True
		
		' Imprimiendo el archivo
		'ActiveSheet.Application.Dialogs(xlDialogPrint).Show
		
		' Muestra el archivo en carpeta para enviar por mail
		Shell "explorer " & rutaViajante, vbNormalFocus
	End If
End Sub

Private Sub Workbook_Activate()
	' Pretegiendo el libraco
	ActiveWorkbook.Protect Password:="Rerda", Structure:=True, Windows:=True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
	' Pretegiendo el libraco
	ActiveWorkbook.Protect Password:="Rerda", Structure:=True, Windows:=True
End Sub

Private Sub Workbook_Open()
	' Pretegiendo el libraco
	ActiveWorkbook.Protect Password:="Rerda", Structure:=True, Windows:=True
	
	' Ocultando una hoja innecesaria
	If Worksheets("Lince").Visible = True Then
		Worksheets("Lince").Visible = False
	End If
	If Worksheets("Datos").Visible = True Then
		Worksheets("Datos").Visible = False
	End If
	
	' Protegiendo el resto de las hojas de cambios innecesarios
	Dim cuenta As Integer
	For cuenta = 1 To Worksheets.Count
		Worksheets(cuenta).Protect
	Next
End Sub
