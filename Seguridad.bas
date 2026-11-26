Attribute VB_Name = "Seguridad"
Option Explicit
Public Const CLAVE_LIBRO As String = "Rerda2026"
Public rutaViajante As String

Sub GenerarRotulo()
    Dim ruta As String
    Dim viajante As String
    Dim fecha As String
    Dim cliente As String
    Dim cp As String
    Dim direccion As String
    Dim telefono As String
    Dim dni As String
    Dim ciudad As String
    Dim provincia As String
    Dim Domicilio As Worksheet
    Dim Sucursal As Worksheet
    Dim Planilla As Workbook
    Dim Proforma As Worksheet
    Dim Datos As Worksheet
    
    Set Planilla = ActiveWorkbook
    Set Domicilio = Planilla.Worksheets("DOMICILIO")
    Set Sucursal = Planilla.Worksheets("SUCURSAL")
    Set Proforma = Planilla.Worksheets("PROFORMA")
    Set Datos = Planilla.Worksheets("DATOS")
    ruta = Planilla.Path & "\Rotulos"
    
    ' A DOMICILIO ========================
    If ActiveSheet Is Domicilio Then
        
        ' El viajante o sucursal
        viajante = UCase(Domicilio.Range("C29").Value)
        If viajante = "" Then
            MsgBox ("¿Y quién lo envía?")
            Domicilio.Range("D29").Activate
            Exit Sub
        Else
            Domicilio.Range("D29").Value = viajante
        End If
        
        fecha = Format(CDate(Domicilio.Range("I29").Value), "yyyy-mm-dd")
        
        ' El cliente
        cliente = UCase(Domicilio.Range("C30").Value)
        If cliente = "" Then
            MsgBox ("Debes completar el cliente")
            Domicilio.Range("C30").Activate
            Exit Sub
        Else
            Domicilio.Range("C30").Value = cliente
        End If
        
        ' CP
        cp = UCase(Domicilio.Range("I31").Value)
        If cp = "" Then
            MsgBox ("Debes completar el código postal")
            Domicilio.Range("I31").Activate
            Exit Sub
        Else
            Domicilio.Range("I31").Value = cp
        End If
        
        ' DNI/CUIT
        dni = UCase(Domicilio.Range("I32").Value)
        If dni = "" Then
            MsgBox ("Debes completar el DNI o el CUIT del cliente")
            Domicilio.Range("I32").Activate
            Exit Sub
        Else
            Domicilio.Range("I32").Value = dni
        End If
        
        ' Teléfono
        telefono = UCase(Domicilio.Range("I33").Value)
        If telefono = "" Then
            MsgBox ("Debes completar el teléfono del cliente")
            Domicilio.Range("I33").Activate
            Exit Sub
        Else
            Domicilio.Range("I33").Value = telefono
        End If
        
        ' PROPIOS DE LA OPCION A DOMICILIO
        ' La Dirección
        direccion = UCase(Domicilio.Range("C31").Value)
        If direccion = "" Then
            MsgBox ("Debes completar la dirección")
            Domicilio.Range("C31").Activate
            Exit Sub
        Else
            Domicilio.Range("C31").Value = direccion
        End If
        
        ' La Ciudad
        ciudad = UCase(Domicilio.Range("C32").Value)
        If ciudad = "" Then
            MsgBox ("Debes completar la ciudad")
            Domicilio.Range("C32").Activate
            Exit Sub
        Else
            Domicilio.Range("C32").Value = ciudad
        End If
        
        ' La Provincia
        provincia = UCase(Domicilio.Range("C33").Value)
        If provincia = "" Then
            MsgBox ("Debes elegir la provincia")
            Domicilio.Range("C33").Activate
            Exit Sub
        Else
            Domicilio.Range("C33").Value = provincia
        End If
        
        ' Dando formato a la página para imprimir
        With Domicilio.PageSetup
            .Orientation = xlPortrait
            .TopMargin = Application.CentimetersToPoints(1)
            .RightMargin = Application.CentimetersToPoints(0.6)
            .BottomMargin = Application.CentimetersToPoints(1)
            .LeftMargin = Application.CentimetersToPoints(0.6)
            .HeaderMargin = Application.CentimetersToPoints(0.2)
            .FooterMargin = Application.CentimetersToPoints(0.2)
            .CenterHorizontally = True
            .Zoom = 115
            .FitToPagesTall = 1
            .FitToPagesWide = 1
            .PaperSize = xlPaperA4
        End With
        
        ' CONTROLANDO PROFORMA
        If provincia = Datos.Range("A24").Value Then
            Call FacProforma(Planilla, Domicilio, cliente, dni, direccion, ciudad, cp, provincia, telefono, viajante)
        End If

        ' Creando las carpetas
        Call crearArchivo(ruta, viajante, cliente)
        Call guardar(Domicilio, rutaViajante, cliente, fecha)

    ' A SUCURSAL =========================
    ElseIf ActiveSheet Is Sucursal Then
        
        ' El viajante o sucursal
        viajante = UCase(Sucursal.Range("C29").Value)
        If viajante = "" Then
            MsgBox ("¿Y quién lo envía?")
            Sucursal.Range("C29").Activate
            Exit Sub
        Else
            Sucursal.Range("C29").Value = viajante
        End If
        
        ' Formato correcto de fecha
        fecha = Format(CDate(Sucursal.Range("J29").Value), "yyyy-mm-dd")
        
        ' El cliente
        cliente = UCase(Sucursal.Range("C30").Value)
        If cliente = "" Then
            MsgBox ("Debes completar el cliente")
            Sucursal.Range("C30").Activate
            Exit Sub
        Else
            Sucursal.Range("C30").Value = cliente
        End If
        
         ' CP
        cp = UCase(Sucursal.Range("J31").Value)
        If cp = "" Then
            MsgBox ("Debes completar el código postal")
            Sucursal.Range("J31").Activate
            Exit Sub
        Else
            Sucursal.Range("J31").Value = cp
        End If
        
        ' DNI/CUIT
        dni = UCase(Sucursal.Range("J32").Value)
        If dni = "" Then
            MsgBox ("Debes completar el DNI o el CUIT del cliente")
            Sucursal.Range("J32").Activate
            Exit Sub
        Else
            Sucursal.Range("J32").Value = dni
        End If
        
        ' Teléfono
        telefono = UCase(Sucursal.Range("J33").Value)
        If telefono = "" Then
            MsgBox ("Debes completar el teléfono del cliente")
            Sucursal.Range("J33").Activate
            Exit Sub
        Else
            Sucursal.Range("J33").Value = telefono
        End If
        
        ' Dirección
        direccion = UCase(Sucursal.Range("C31").Value)
        
        ' Ciudad
        ciudad = UCase(Sucursal.Range("C32").Value)
        
        ' Provincia
        provincia = UCase(Sucursal.Range("C33").Value)
        
        ' Formato correcto de impresión
        With Sucursal.PageSetup
            .Orientation = xlPortrait
            .TopMargin = Application.CentimetersToPoints(1)
            .RightMargin = Application.CentimetersToPoints(0.6)
            .BottomMargin = Application.CentimetersToPoints(1)
            .LeftMargin = Application.CentimetersToPoints(0.6)
            .HeaderMargin = Application.CentimetersToPoints(0.2)
            .FooterMargin = Application.CentimetersToPoints(0.2)
            .CenterHorizontally = True
            .Zoom = 115
            .FitToPagesTall = 1
            .FitToPagesWide = 1
            .PaperSize = xlPaperA4
        End With
        
        ' CONTROLANDO PROFORMA
        If provincia = Datos.Range("A24").Value Then
            Call FacProforma(Planilla, Sucursal, cliente, dni, direccion, ciudad, cp, provincia, telefono, viajante)
        End If
        
        ' Creando las carpetas
        Call crearArchivo(ruta, viajante, cliente)
        Call guardar(Sucursal, rutaViajante, cliente, fecha)
        
    End If
    
End Sub

' CHEQUEA SI CORRESPONDE O NO LA F. PROFORMA
Function FacProforma(Planilla As Workbook, origen As Worksheet, apellidoNombre As String, dni As String, direccion As String, ciudad As String, codigoPostal As String, provincia As String, telefono As String, viajante As String)
    
    ' Variables a utilizar
    Dim acumulador As Byte
    Dim cantidad As Byte
    Dim precio As Double
    Dim sku As String
    Dim color As String
    Dim talle As String
    Dim i As Byte
    
    ' Borra información previa
    With Planilla.Worksheets("PROFORMA")
        .Activate
        .Range("I1").ClearContents
        .Range("A22:D46").ClearContents
        .Range("H22:H46").ClearContents
        .Range("I7:I15").ClearContents
        .Range("I18:I19").ClearContents
    End With
    
    
    If provincia = Planilla.Worksheets("DATOS").Range("A24").Value Then
        ' Hace proforma
        With Planilla.Worksheets("PROFORMA")
            .Cells(1, 9).Value = UCase(viajante)
            .Cells(7, 9).Value = UCase(apellidoNombre)
            .Cells(9, 9).Value = dni
            .Cells(10, 9).Value = UCase(direccion)
            .Cells(12, 9).Value = UCase(ciudad)
            .Cells(13, 9).Value = codigoPostal
            .Cells(14, 9).Value = UCase(provincia)
            .Cells(18, 9).Value = "'" & telefono
        End With
        
        acumulador = 0
        
        ' Bucle que recorre la venta
        ' Cantidad de filas de productos: 14.
        ' Comienza en la fila 36 y termina en la 49
        For i = 36 To 49
            ' Toma datos de la "Planilla"
            With origen
                sku = .Cells(i, 1).Value
                color = UCase(.Cells(i, 7).Value)
                talle = UCase(.Cells(i, 8).Value)
                cantidad = .Cells(i, 9).Value
            End With
                   
            ' Vuelca info en la Proforma
            With Planilla.Worksheets("PROFORMA")
                ' Copia el SKU
                .Cells(22 + acumulador, 2).Value = sku
            
                ' Copia el talle
                .Cells(22 + acumulador, 4).Value = talle
            
                ' Copia el color
                .Cells(22 + acumulador, 3).Value = color
            
                ' Copia la cantidad
                If cantidad > 0 Then
                    .Cells(22 + acumulador, 1).Value = cantidad
                End If
            End With
                        
            ' Incrementa el contador
            acumulador = acumulador + 1
            
        Next i
        
        ' Indica que se complete información de precios
        With Planilla.Worksheets("PROFORMA")
            .Activate
            MsgBox ("Ahora completá los precios unitarios de los productos.")
            .Range("H22").Activate
        End With
    End If
End Function


Sub imprimeProforma()
    ' GUARDA EN PDF E IMPRIME LA FACTURA PROFORMA
    Dim Proforma As Workbook
    Dim i As Byte
    Dim fecha As String
    Dim ruta As String
    Dim viajante As String
    Dim cliente As String
    
    Set Proforma = ActiveWorkbook
    
    ' Dando forma a la impresión
    With Proforma.Worksheets("PROFORMA").PageSetup
        .Orientation = xlPortrait
        .TopMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(0.6)
        .BottomMargin = Application.CentimetersToPoints(1)
        .LeftMargin = Application.CentimetersToPoints(0.6)
        .HeaderMargin = Application.CentimetersToPoints(0.8)
        .FooterMargin = Application.CentimetersToPoints(0.8)
        .CenterHorizontally = True
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
        '.CenterHeader = "&B&20&F"
        .PaperSize = xlPaperA4
    End With
    
    With Proforma.Worksheets("PROFORMA")
        ' Posicionarse para los precios en dólares
        .Cells(22, 1).Activate
        
        ' Asignando valores
        viajante = .Range("I1").Value
        fecha = Format(CDate(.Range("I4").Value), "yyyy-mm-dd")
        cliente = UCase(.Range("I7").Value)
    
    End With
    ruta = Proforma.Path & "\Rotulos"
    
    
    ' Recorre los precios unitarios para que estén completos.
    Do While ActiveCell.Value <> ""
        If ActiveCell.Offset(0, 7).Value = "" Then
            MsgBox "Te faltó un precio unitario"
            ActiveCell.Offset(0, 7).Activate
            Exit Sub
        End If
        ActiveCell.Offset(1, 0).Activate
    Loop
        
    ' Imprimie la proforma
    
    Call crearArchivo(ruta, viajante, cliente)
    Call guardar(Proforma.Worksheets("PROFORMA"), rutaViajante, cliente & " - Factura Proforma", fecha)
End Sub


' CHEQUEA QUE EXISTA LA CARPETA DONDE SE GUARDA, SI NO, LA CREA.
Function crearArchivo(ruta, viajante, cliente)
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
Function guardar(Planilla As Worksheet, rutaViajante, cliente, fecha)

    ' Muestra en qué carpeta se guarda
    MsgBox ("Se guardó una copia PDF en " & rutaViajante & "\" & fecha & ". " & cliente & ".pdf")
    
    ' Cambio de nomenclatura en el nombrado ascendente de los archivos.
    Planilla.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        rutaViajante & "\" & fecha & ". " & cliente & ".pdf", Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        True
    
    ' Muestra el archivo en carpeta para enviar por mail o imprimir.
    Shell "explorer " & rutaViajante, vbNormalFocus
End Function

Sub Proteger()
    Dim Libraco As Workbook
    Dim ws As Worksheet
    
    Set Libraco = ActiveWorkbook
    
    ' Proteger cada hoja si está desprotegida
    For Each ws In Libraco.Worksheets
        If Not ws.ProtectContents Then
            ws.Protect Password:=CLAVE_LIBRO
        End If
    Next ws
    
    ' Proteger la estructura del libro si está desprotegida
    If Not Libraco.ProtectStructure Then
        Libraco.Protect Password:=CLAVE_LIBRO, Structure:=True, Windows:=False
    End If
End Sub

Sub Desproteger()
    Dim Libraco As Workbook
    Dim ws As Worksheet
    
    Set Libraco = ActiveWorkbook
    
    ' Desproteger la estructura del libro si está protegida
    If Libraco.ProtectStructure Then
        Libraco.Unprotect Password:=CLAVE_LIBRO
    End If
    
    ' Desproteger cada hoja si está protegida
    For Each ws In Libraco.Worksheets
        If ws.ProtectContents Then
            ws.Unprotect Password:=CLAVE_LIBRO
        End If
    Next ws
End Sub

