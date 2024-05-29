' ============================================================
' ==================== MODULO 1 ==============================
' ============================================================
Option Explicit
' Definiendo unas variables globales a usar
Dim cp As String
Dim codigoNis As String
Dim dni As String
Dim ruta As String
Dim rutaViajante As String
Dim viajante As String
Dim cliente As String
Dim fecha As String
Dim domicilio As String
Sub Generar_Rotulo()
' ============================================================
' GENERA UN ROTULO Y PLANILLA PARA IMPRIMIR LA VENTA DE UN
' VIAJANTE O DE UNA SUCURSAL.
' ============================================================


' Asigna valores de acuerdo a datos de la planilla actual
ruta = ThisWorkbook.Path & "\Ventas de Viajantes"
viajante = UCase(ActiveSheet.Range("X1").Value)
'fecha = Day(Date) & "-" & Month(Date) & "-" & Year(Date)

' Nuevo formato para dar orden ascendente a los archivos
fecha = Format(Date, "yyyy-mm-dd")

dni = ActiveSheet.Range("X6").Value
cliente = UCase(ActiveSheet.Range("X2").Value)
cp = ActiveSheet.Range("X5").Value
domicilio = Sheets("A Domicilio").Range("X3").Value

' Controla si la carpeta existe, de lo contrario, la crea en local
Call crearArchivo(ruta, viajante, fecha, cliente, cp)
    
' Controla si se determinó el vendedor, viajente o sucursal
If viajante = "" Then
    MsgBox ("¿Y tu nombre como vendedor, viajante o sucursal?")
    ActiveSheet.Range("X1").Activate

' Controla apellido, nombre o razón social del cliente
ElseIf ActiveSheet.Range("X2").Value = "" Then
    MsgBox ("¿Y el Apellido/Nombre o razón social del cliente?")
    ActiveSheet.Range("X2").Activate

' Controla datos de dirección para el caso de "ENVIO A DOMICILIO".
ElseIf ActiveSheet.Name = "A Domicilio" And ActiveSheet.Range("X3").Value = "" Then
    MsgBox ("¿Y la dirección de destino? Calle, Nº, piso, dpto, etc...")
    Sheets("A Domicilio").Range("X3").Activate

' Controla que el código postal exista y sea correcto.
ElseIf ActiveSheet.Range("X5").Value = "" Then
    MsgBox ("¿Y el código postal?")
    ActiveSheet.Range("X5").Activate

' Controla el DNI/CUIT del cliente
ElseIf ActiveSheet.Range("X6").Value = "" Then
    MsgBox ("¿Y el DNI o CUIL/CUIT del cliente?")
    ActiveSheet.Range("X6").Activate

' Controla el teléfono de contacto del cliente
ElseIf ActiveSheet.Range("X7").Value = "" Then
    MsgBox ("¿Y el teléfono o celular del cliente?")
    ActiveSheet.Range("X7").Activate

' Controla la ciudad o pueblo de destino
ElseIf ActiveSheet.Range("X8").Value = "" Then
    MsgBox ("¿Y la ciudad o pueblo?")
    ActiveSheet.Range("X8").Activate

' Controla la Provincia de destino
ElseIf ActiveSheet.Range("X9").Value = "" Then
    MsgBox ("¿Y la provincia?")
    ActiveSheet.Range("X9").Activate

' Controla que sea NIS para el caso de retiro en sucursal
ElseIf ActiveSheet.Name = "A Sucursal" Or ActiveSheet.Name = "Pago en Destino" Then
    domicilio = "Sucursal Correo Argentino NIS " & Application.VLookup(ActiveSheet.Range("X5"), Sheets("Sucursales").Range("A3:B4000"), 2, False)
    
    ' Muestra dónde se guarda el archivo y lo abre
    Call guardar(rutaViajante, cliente, fecha)
    
    ' Genera, si corresponde; la factura proforma
    Call proforma(cliente, dni, domicilio, UCase(ActiveSheet.Range("X9").Value), cp, UCase(ActiveSheet.Range("X8").Value), UCase(ActiveSheet.Range("X7").Value))

Else
    ' Muestra dónde se guarda el archivo y lo abre
    Call guardar(rutaViajante, cliente, fecha)
    
    ' Genera, si corresponde; la factura proforma
    Call proforma(cliente, dni, domicilio, UCase(ActiveSheet.Range("X9").Value), cp, UCase(ActiveSheet.Range("X8").Value), UCase(ActiveSheet.Range("X7").Value))
End If

End Sub

' FORMATEA LA PLANILLA PARA VISUALIZACION Y MODO IMPRESION
Function darFormato()
    ' Formato a cada planilla en cuestión
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
    End With
    
    ' FACTURA PROFORMA ==========================
    With Sheets("Proforma").PageSetup
        .Orientation = xlPortrait
        .TopMargin = Application.CentimetersToPoints(1.9)
        .RightMargin = Application.CentimetersToPoints(0.6)
        .BottomMargin = Application.CentimetersToPoints(1.9)
        .LeftMargin = Application.CentimetersToPoints(0.6)
        .HeaderMargin = Application.CentimetersToPoints(0.8)
        .FooterMargin = Application.CentimetersToPoints(0.8)
        .CenterHorizontally = True
        .PaperSize = xlPaperA4
    End With
    
End Function


' CHEQUEA QUE EXISTA LA CARPETA DONDE SE GUARDA, SI NO, LA CREA.
Function crearArchivo(ruta, viajante, fecha, cliente, cp)
    If Dir(ruta, vbDirectory) = "" Then
        MkDir (ruta)
    End If
    
    ' Controla si existe la carpeta del vendedor/viajante/sucursal
    rutaViajante = ruta & "\" & viajante
    If Dir(rutaViajante, vbDirectory) = "" Then
        MkDir (rutaViajante)
    End If
End Function

' GUARDA EL ARCHIVO CON NOMBRE DE CLIENTE Y LA FECHA DE CREACION.
' TAMBIEN MUESTRA DONDE GUARDA Y LO ABRE PARA IMPRIMIR.
Function guardar(rutaViajante, cliente, fecha)

    ' Muestra en qué carpeta se guarda
    MsgBox ("Se guardó una copia PDF en " & rutaViajante & "\" & fecha & ". " & cliente & ".pdf")
    
    ' Cambio de nomenclatura en el nombrado ascendente de los archivos.
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        rutaViajante & "\" & fecha & ". " & cliente & ".pdf", Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        True
    
    ' Muestra el archivo en carpeta para enviar por mail o imprimir.
    Shell "explorer " & rutaViajante, vbNormalFocus
End Function

' CHEQUEA SI CORRESPONDE O NO LA F. PROFORMA
Function proforma(apellidoNombre, dni, direccion, provincia, codigoPostal, ciudad, telefono)
    ' Variables a utilizar
    Dim acumulador As Byte
    Dim cantidad As Byte
    Dim precio As Double
    Dim sku As String
    Dim color As String
    Dim talle As String
    
    ' Borra información previa
    With Sheets("Proforma")
        .Range("A22:D46").ClearContents
        .Range("H22:H46").ClearContents
        .Range("I7:I15").ClearContents
        .Range("I18:I19").ClearContents
    End With
    
    
    If provincia = "TIERRA DEL FUEGO" Or provincia = "Tierra del Fuego" Then
        ' Hace proforma
        With Sheets("Proforma")
            .Cells(7, 9).Value = UCase(apellidoNombre)
            .Cells(9, 9).Value = dni
            .Cells(10, 9).Value = direccion
            .Cells(12, 9).Value = UCase(ciudad)
            .Cells(13, 9).Value = codigoPostal
            .Cells(14, 9).Value = UCase(provincia)
            .Cells(18, 9).Value = "'" & telefono
        End With
        acumulador = 0
        
        ' Bucle que recorre la venta
        ActiveSheet.Activate
        Range("W12").Activate
        
        Do While ActiveCell.Value <> ""
        
            ' Toma datos de la "Planilla"
            With ActiveSheet
                sku = ActiveCell.Value
                cantidad = ActiveCell.Offset(0, 4).Value
                color = ActiveCell.Offset(0, 2).Value
                talle = ActiveCell.Offset(0, 3).Value
            End With
                   
            ' Vuelca info en la Proforma
            With Sheets("Proforma")
                ' Copia el SKU
                .Cells(22 + acumulador, 2).Value = sku
            
                ' Copia el talle
                .Cells(22 + acumulador, 4).Value = talle
            
                ' Copia el color
                .Cells(22 + acumulador, 3).Value = color
            
                ' Copia la cantidad
                .Cells(22 + acumulador, 1).Value = cantidad
                
            End With
                        
            ' Incrementa el contador
            acumulador = acumulador + 1
            
            ' Salta a la siguiente celda de abajo
            ActiveCell.Offset(1, 0).Activate
            
            ' Sale cuando no hay más SKU
            If ActiveCell.Value = "" Then
                Exit Do
            End If
        Loop
        
        ' Indica que se complete información de precios
        Sheets("Proforma").Activate
        MsgBox ("Ahora completá los precios unitarios de los productos.")
        Range("H22").Activate
    End If
End Function


Sub imprimeProforma()
    ' GUARDA EN PDF E IMPRIME LA FACTURA PROFORMA
    Dim permitido As Boolean
    Range("A22").Activate
    
    ' Recorre los precios unitarios para que estén completos.
    Do While ActiveCell.Value <> ""
        If ActiveCell.Offset(0, 7).Value = "" Then
            MsgBox "Te faltó un precio unitario"
            ActiveCell.Offset(0, 7).Activate
            Exit Sub
        End If
        ActiveCell.Offset(1, 0).Activate
    Loop
    
    ' Controla el código NIS que esté completo
    If Sheets("Proforma").Range("I10").Value = "" Then
        MsgBox "Faltó poner código NIS en el domicilio"
        Sheets("Proforma").Range("I10").Activate
        Exit Sub
    End If
        
    ' Imprimie la proforma
    Call guardar(rutaViajante, cliente & " - Factura Proforma", fecha)
End Sub


' ============================================================
' ==================== ThisWorkbook ==========================
' ============================================================
Option Explicit
dim clave as String

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Cancel = False
    ThisWorkbook.Save
End Sub

Private Sub Workbook_Open()
    clave = "¿Qué te importa?"
    
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
    ThisWorkbook.Protect (clave)

End Sub