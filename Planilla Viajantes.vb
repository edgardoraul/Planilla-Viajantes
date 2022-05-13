Option Explicit
Sub Generar_Rotulo()
' ============================================================
' GENERA UN ROTULO Y PLANILLA PARA IMPRIMIR VENTA DE UN
' VIAJANTE O SUCURSASL. CONTROLA QUE LAS SUBCARPETAS SE HAYAN
' CREADO. SI NO, LAS CREA. TAMBIÉN GUARDA LOS ARCHIVOS CON
' NOMBRE DE LOS CLIENTES Y FECHA, CADA RESPECTIVA SUBCARPETA.
' ============================================================
    ' Pretegiendo el libraco
    ThisWorkbook.Protect
    
    ' Ocultando una hoja innecesaria
    If Worksheets("Lince").Visible = True Then
        Worksheets("Lince").Visible = False
    End If
    
    ' Protegiendo el resto de las hojas de cambios innecesarios
    Dim cuenta As Integer
    For cuenta = 1 To Worksheets.Count
        Worksheets(cuenta).Protect
    Next
    
    ' Definiendo variables a usar
    Dim ruta As String
    Dim rutaViajante As String
    Dim viajante As String
    Dim cliente As String
    Dim fecha As String
    Dim cp As String
    Dim codigoNis As String
    
    
    ' Asigando los valores de acuerdo a la planilla donde esté parado
    ruta = ThisWorkbook.Path & "\Ventas de Viajantes"
    viajante = ActiveSheet.Range("X1").Value
    fecha = Day(Date) & "-" & Month(Date) & "-" & Year(Date)
    cliente = ActiveSheet.Range("X2").Value
    cp = ActiveSheet.Range("X5").Value
    
    
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
    
    
    ' Controlando al vendedor o viajante
    If viajante = "" Then
        MsgBox ("¿Y tu nombre como vendedor, viajante o sucursal?")
        ActiveSheet.Range("X1").Select
    
    ' Controlando apellido, nombre razón social del cliente
    ElseIf ActiveSheet.Range("X2").Value = "" Then
        MsgBox ("¿Y el Apellido/Nombre o razón social del cliente?")
        ActiveSheet.Range("X2").Select
    
    ' Controlando la dirección
    ElseIf Sheets("A Domicilio").Range("X3").Value = "" Then
        MsgBox ("¿Y la dirección de destino? Calle, Nº, piso, dpto, etc...")
        ActiveSheet.Range("X3").Select
        
    ' Controlando el código postal
    ElseIf ActiveSheet.Range("X5").Value = "" Then
        MsgBox ("¿Y el código postal?")
        ActiveSheet.Range("X5").Select
        
    ' Controlando el código postal
    ElseIf ActiveSheet.Name = "A Sucursal" And Sheets("A Sucursal").Range("X5").Value = "" Then
        ' Validando si lo encuentra
        If codigoNis = "" Then
            ' Todo salió mal
            Sheets("A Sucursal").Range("X5").Select
            codigoNis = Sheets("Sucursales").Range("$F$4:$F$5000").Find(What:=cp, LookIn:=xlValues, LookAt:=xlPart).Offset(0, -4)
            MsgBox ("El código postal " & cp & " no existe. Intentá con otro. ")
        End If
    
    ' Controlando el DNI/CUIT
    ElseIf ActiveSheet.Range("X6").Value = "" Then
        MsgBox ("¿Y el DNI o CUIL/CUIT del cliente?")
        ActiveSheet.Range("X6").Select
    
    ' Controlando el teléfono de contacto
    ElseIf ActiveSheet.Range("X7").Value = "" Then
        MsgBox ("¿Y el teléfono o celular del cliente?")
        ActiveSheet.Range("X7").Select
    
    ' Controlando la ciudad o pueblo
    ElseIf Sheets("A Domicilio").Range("X8").Value = "" Then
        MsgBox ("¿Y la ciudad o pueblo?")
        ActiveSheet.Range("X8").Select
    
    Else
    
        ' Guardando el archivo
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            rutaViajante & "\" & cliente & ". " & fecha & ".pdf", Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
            False
        
        ' Mostrar dónde se guardó el archivo
        MsgBox ("Se guardó una copia PDF en " & rutaViajante & "\" & cliente & ". " & fecha & ".pdf")
        
        ' Imprimiendo el archivo
        ActiveSheet.Application.Dialogs(xlDialogPrint).Show
    End If
    
End Sub
