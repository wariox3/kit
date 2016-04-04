Attribute VB_Name = "ModFuncionesPropias"
Option Explicit

Sub IniProg(Ini As Long, Fin As Long)
On Error GoTo MiError
  Principal.PgsPrincipal.Width = Principal.PicMensajes.Width
  Principal.PgsPrincipal.Min = 1
  Principal.PgsPrincipal.Max = Fin
MiError:
End Sub
Sub Prog(Valor)
On Error GoTo MiError
  Principal.PgsPrincipal.value = Valor
MiError:
End Sub
Sub FinProg()
  Principal.PgsPrincipal.Width = 0
End Sub
Sub MsgTit(Mensaje As String)
  Principal.LblMensaje = Mensaje
  Principal.TmPrincipal.Enabled = True
End Sub


Function DevListaPrecios(IdCiudad As Long, IdProducto As Long, IdListaPrecios As Integer, IdListaGeneral As Integer, ConGeneral As Boolean) As TipListaPrecios
    AbrirRecorset rstUniversal, "Select*from ListasPreciosCiudades where idListaPrecios=" & IdListaPrecios & " and IdCiudad=" & IdCiudad & " and IdProducto=" & IdProducto, CnnPrincipal, adOpenKeyset, adLockReadOnly
    If rstUniversal.EOF = True Then
      MsgBox "No hay precios para esta ciudad con este cliente en esta lista de precios", vbCritical
      If ConGeneral = True Then
        AbrirRecorset rstUniversal, "Select*from ListasPreciosCiudades where idListaPrecios=" & IdListaGeneral & " and IdCiudad=" & IdCiudad & " and IdProducto=" & IdProducto, CnnPrincipal, adOpenKeyset, adLockReadOnly
        If rstUniversal.EOF = False Then
          If MsgBox("¿Se ha encontrado un precio en la lista de precios general, desea aplicarle este precio a la liquidacion?", vbQuestion + vbYesNo) = vbYes Then
            DevListaPrecios.Devuelve = True
            DevListaPrecios.VrKilo = rstUniversal!VrKilo
            DevListaPrecios.KTope = rstUniversal!KTope
            DevListaPrecios.VrKTope = rstUniversal!VrKTope
            DevListaPrecios.VrKdicional = rstUniversal!VrKAdicional
            DevListaPrecios.VrUnidad = rstUniversal!VrUnidad
            DevListaPrecios.Minimos = rstUniversal!Minimos
          Else
            DevListaPrecios.Devuelve = False
          End If
        End If
      Else
        DevListaPrecios.Devuelve = False
      End If
    Else
      DevListaPrecios.Devuelve = True
      DevListaPrecios.VrKilo = rstUniversal!VrKilo
      DevListaPrecios.KTope = rstUniversal!KTope
      DevListaPrecios.VrKTope = rstUniversal!VrKTope
      DevListaPrecios.VrKdicional = rstUniversal!VrKAdicional
      DevListaPrecios.VrUnidad = rstUniversal!VrUnidad
      DevListaPrecios.Minimos = rstUniversal!Minimos
      
    End If
    CerrarRecorset rstUniversal
End Function
Function DevEstadoDespacho(Estado As String) As String
  Select Case Estado
    Case "D"
      DevEstadoDespacho = "DIGITADO"
    Case "V"
      DevEstadoDespacho = "VIAJANDO"
    Case "I"
      DevEstadoDespacho = "IMPRESO"
    Case "A"
      DevEstadoDespacho = "ANULADO"
    Case "G"
      DevEstadoDespacho = "DESCARGADO"
    Case "U"
      DevEstadoDespacho = "REPARTO"
    Case "E"
      DevEstadoDespacho = "DESEMBARCADA"
    Case "P"
      DevEstadoDespacho = "PLANILLANDO"
  End Select
End Function
Public Function DevNombreDatosBasicos(Id As String) As String
  AbrirRecorset rstUniversal, "SELECT IdTercero, Nombre, Apellido1, Apellido2 From Terceros Where IdTercero ='" & Id & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = False Then
    DevNombreDatosBasicos = rstUniversal.Fields("Nombre") & " " & rstUniversal.Fields("Apellido1") & " " & rstUniversal.Fields("Apellido2")
  End If
  CerrarRecorset rstUniversal
End Function

Public Function ComprobarEstadoSel(Guia As Long, Tipo As Integer) As Boolean
  Dim rstComprobar As New ADODB.Recordset
  Dim Campo As String
  rstComprobar.CursorLocation = adUseClient
  ComprobarEstadoSel = False
  rstComprobar.Open "Select guia, Entregada, Descargada, Despachada, Anulada, Facturada from Guias where Guia=" & Guia, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstComprobar.EOF = False Then
    If Tipo = 0 Then
      If Val(rstComprobar.Fields(1)) = 1 And Val(rstComprobar.Fields(2)) = 1 And Val(rstComprobar.Fields(3)) = 1 Then
        ComprobarEstadoSel = True
      Else
        ComprobarEstadoSel = False
      End If
    Else
      If Val(rstComprobar.Fields(Tipo)) = 1 Then
        ComprobarEstadoSel = True
      Else
        ComprobarEstadoSel = False
      End If
    End If
  End If
  rstComprobar.Close
  Set rstComprobar = Nothing
End Function

Public Function ComprobarExGuia(Guia As Long, Tipo As Integer) As Boolean
  Dim rstComprobar As New ADODB.Recordset
  rstComprobar.CursorLocation = adUseClient
  ComprobarExGuia = False
  rstComprobar.Open "SELECT Guia FROM guias WHERE Guia = " & Guia & " AND GuiaTipo = " & Tipo, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstComprobar.EOF = False Then
    ComprobarExGuia = True
  End If
  rstComprobar.Close
  Set rstComprobar = Nothing
End Function
Public Function ComprobarExGuiaGeneral(Guia As Long) As Boolean
  Dim rstComprobar As New ADODB.Recordset
  rstComprobar.CursorLocation = adUseClient
  ComprobarExGuiaGeneral = False
  rstComprobar.Open "SELECT Guia FROM guias WHERE Guia = " & Guia, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstComprobar.EOF = False Then
    ComprobarExGuiaGeneral = True
  End If
  rstComprobar.Close
  Set rstComprobar = Nothing
End Function

Public Function ComprobarEstado(Guia As Long) As String
  Dim rstComprobar As New ADODB.Recordset
  rstComprobar.CursorLocation = adUseClient
  rstComprobar.Open "Select guia, Estado from Guias where Guia=" & Guia, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstComprobar.EOF = False Then
    ComprobarEstado = rstComprobar.Fields("Estado")
  End If
  rstComprobar.Close
  Set rstComprobar = Nothing
End Function


Public Function DevMonSQL(Valor As String) As String
  Dim Tem As String, I As Integer
  For I = 1 To Len(Valor)
    If Mid(Valor, I, 1) <> "." Then
      If Mid(Valor, I, 1) = "," Then
        Tem = Tem & "."
      Else
        Tem = Tem & Mid(Valor, I, 1)
      End If
    End If
  Next
  DevMonSQL = Tem
End Function

Sub BloquearMenu()
  Principal.MnuArchivo.Enabled = False
  Principal.MnuBuscar.Enabled = False
  Principal.MnuMovimiento.Enabled = False
  Principal.MnuAplicar.Enabled = False
  Principal.MnuHerramientas.Enabled = False
  Principal.MnuComplementos.Enabled = False
  Principal.MnuRutinas.Enabled = False
  Principal.MnuInfRepLis.Enabled = False
  Principal.MnuAyuda.Enabled = False
End Sub
Sub DesBloquearMenu()
  Principal.MnuArchivo.Enabled = True
  Principal.MnuBuscar.Enabled = True
  Principal.MnuMovimiento.Enabled = True
  Principal.MnuAplicar.Enabled = True
  Principal.MnuHerramientas.Enabled = True
  Principal.MnuComplementos.Enabled = True
  Principal.MnuRutinas.Enabled = True
  Principal.MnuInfRepLis.Enabled = True
  Principal.MnuAyuda.Enabled = True
End Sub

Public Sub GenerarReciboCaja(Guia As Long, TpFlete As Integer, TpManejo As Integer, Flete As Double, Manejo As Double)
    Dim FleteF As Double, ManejoF As Double
    If TpFlete = 1 Then
        FleteF = Flete
    End If
    
    If TpManejo = 1 Then
        ManejoF = Manejo
    End If
    
    AbrirRecorset rstUniversal, "Insert into recibos (NroRecibo, FechaRecibo, GuiaRecibo, Flete, Manejo) values (" & SacarConsecutivo("Recibos", CnnPrincipal) & ", now(), " & Guia & ", " & FleteF & ", " & ManejoF & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
End Sub

Public Function DevDocSinCeros(Doc As String) As String
  Dim salida As String
  Dim I As Integer
  Dim N As Boolean
  N = False
  I = 1
  Do While I <= Len(Doc)
    
    If N = False And (Mid(Doc, I, 1) <> "0" And Mid(Doc, I, 1) <> "'") Then
      N = True
    End If
    
    If N = False And (Mid(Doc, I, 1) = "0" Or Mid(Doc, I, 1) = "'") Then
      salida = salida
    Else
      salida = salida & Mid(Doc, I, 1)
    End If
    I = I + 1
  Loop
  DevDocSinCeros = salida
End Function

Public Sub ExportarExcel(rstTemp As ADODB.Recordset)
  Dim RutaSalida As String
  Dim o_Excel     As Object
  Dim o_Libro     As Object
  Dim o_Hoja      As Object
  Dim Fila        As Long
  Dim Columna     As Long
  
On Error GoTo Error_Handler

  Principal.CDExa.DialogTitle = "Guardar como"
  Principal.CDExa.Filter = "Archivo Excel|*.xls"
  Principal.CDExa.ShowSave
  If Principal.CDExa.FileName <> "" Then
    RutaSalida = Principal.CDExa.FileName
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
    For Columna = 1 To rstTemp.Fields.Count
      o_Hoja.Cells(1, Columna).value = rstTemp.Fields(Columna - 1).Name
    Next
    
    With rstTemp
        For Fila = 2 To .RecordCount + 1
            For Columna = 0 To .Fields.Count - 1
                o_Hoja.Cells(Fila, Columna + 1).value = .Fields(Columna).value
            Next
            .MoveNext
        Next
    End With
    o_Libro.Close True, RutaSalida
    o_Excel.Quit
    
  End If
  Exit Sub
Error_Handler:
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
        
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Sub

Public Function EnviarCorreo(Guia As Double, Asunto As String, Mensaje As String, mail As String, oMail As Object) As Boolean
    On Error GoTo ErrControl
    Set oMail = New clsCDOmail
    Dim UsuarioEmail As String
    Dim ClaveEmail As String
    Dim RemiteEmail As String
    Dim rstRegistro As New ADODB.Recordset
    Dim rstUsuario As New ADODB.Recordset
    Dim rstConfiguracion As New ADODB.Recordset
    rstUsuario.CursorLocation = adUseClient
    rstConfiguracion.CursorLocation = adUseClient
    
    AbrirRecorset rstConfiguracion, "SELECT configuracion.* FROM configuracion WHERE Codigo = 1", CnnPrincipal, adOpenDynamic, adLockOptimistic
    AbrirRecorset rstUsuario, "SELECT usuarios.* FROM usuarios WHERE IDUsuario = " & CodUsuarioActivo, CnnPrincipal, adOpenDynamic, adLockOptimistic
    UsuarioEmail = rstUsuario.Fields("UsuarioMail")
    ClaveEmail = rstUsuario.Fields("ClaveMail")
    RemiteEmail = rstUsuario.Fields("RemiteMail")
    Set rstUsuario = Nothing
    With oMail
         'Por defecto estaba autentificacion true y ssl false
         'datos para enviar
        .servidor = rstConfiguracion!ServidorCorreo & ""
        .puerto = rstConfiguracion!puerto & ""
        If Val(rstConfiguracion!UsaAutenticacion) = 1 Then
          .UseAuntentificacion = True
        Else
          .UseAuntentificacion = False
        End If
        If Val(rstConfiguracion!UsaSSL) = 1 Then
          .ssl = True
        Else
          .ssl = False
        End If
        .Usuario = UsuarioEmail
        .PassWord = ClaveEmail
        
        .Asunto = Asunto
        '.Adjunto = "c:\archivo.zip"
        .de = RemiteEmail
        .para = mail
        .Mensaje = Mensaje
        
        .Enviar_Backup ' manda el mail
    End With
ErrControl:
    If Err.Number <> 0 Then
      MsgBox "Ocurrio un error en el envio" & Err.Description
    Else
      AbrirRecorset rstRegistro, "INSERT INTO registro_envios_email (Guia, FechaEnvio, MailEnvio, MailDestino, Usuario) VALUES (" & Guia & ", now(), '" & UsuarioEmail & "', '" & mail & "', " & CodUsuarioActivo & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
    End If
End Function

Public Function Rellenar(Dato As String, Tamaño As Integer, Caracter As String, Posicion As Byte) As String
  FufuSt = ""
  If Len(Dato) < Tamaño Then
    For FufuLo = 1 To Tamaño - Len(Dato)
      FufuSt = FufuSt & Caracter
    Next
    If Posicion = 1 Then
      Rellenar = FufuSt & Dato
    Else
      Rellenar = Dato & FufuSt
    End If
  End If
End Function

Public Sub establecerPapel()
On Error GoTo ErrorImpresora
  Printer.PaperSize = 1
ErrorImpresora:
  'If Err.Number = 380 Then
  '  MsgBox "Error " & Err.Number & " no esta configurada la impresora en el equipo", vbCritical
  'Else
  '  MsgBox "Error controlado " & Err.Description
  'End If
End Sub

Public Sub ExportarGuiaFactura(longGuia As Long)
  Dim rstCuentasCobrar As New ADODB.Recordset
  rstCuentasCobrar.CursorLocation = adUseClient
  Dim rstGuias As New ADODB.Recordset
  Dim douTotal As Double
  AbrirRecorset rstGuias, "SELECT guias.*, terceros.IdAsesor, ciudades.NmCiudad FROM guias LEFT JOIN terceros ON guias.Cuenta = terceros.IDTercero LEFT JOIN ciudades ON guias.IdCiuDestino = ciudades.IdCiudad WHERE ExportadaCartera = 0 AND Guia = " & longGuia, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstGuias.RecordCount > 0 Then
    If Val(rstGuias!GuiFac) = 1 Then
      douTotal = rstGuias!VrFlete + rstGuias!VrManejo
      AbrirRecorset rstCuentasCobrar, "INSERT INTO cuentas_cobrar(NroDocumento, TipoFactura, FechaDoc, FhVence, IdTercero, Total, Saldo, VrFlete, VrManejo, GuiaFactura, IdAsesor, IdPO, Soporte) VALUES (" & rstGuias!Guia & ", " & rstGuias!GuiaTipo & ", '" & Format(rstGuias!FhEntradaBodega, "yyyy-mm-dd") & "', '" & Format(rstGuias!FhEntradaBodega, "yyyy-mm-dd") & "', '" & rstGuias!Cuenta & "', " & douTotal & ", " & douTotal & ", " & rstGuias!VrFlete & ", " & rstGuias!VrManejo & ", 1, " & rstGuias!IdAsesor & ", " & rstGuias!COIng & ", '" & rstGuias!NmCiudad & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
      AbrirRecorset rstUniversal, "INSERT INTO facturas_venta (Numero, TipoFactura, Fecha, FhVence, IdTercero, Plazo, Total, VrFlete, VrManejo, IdPO, IdAsesor) VALUES (" & rstGuias!Guia & ", " & rstGuias!GuiaTipo & ", '" & Format(rstGuias!FhEntradaBodega, "yyyy-mm-dd") & "', '" & Format(rstGuias!FhEntradaBodega, "yyyy-mm-dd") & "', '" & rstGuias!Cuenta & "', 0, " & douTotal & ", " & rstGuias!VrFlete & ", " & rstGuias!VrManejo & ", " & rstGuias!COIng & ", " & rstGuias!IdAsesor & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
      AbrirRecorset rstUniversal, "UPDATE Guias SET ExportadaCartera = 1 WHERE Guia = " & longGuia, CnnPrincipal, adOpenDynamic, adLockOptimistic
    End If
  Else
    MsgBox "No se encuentra la guia " & longGuia & " o ya fue exportada", vbCritical
  End If
End Sub

Public Function DevFechaEspecial() As Date
  AbrirRecorset rstUniversal, "SELECT FechaAfectada, HorasAfectacion, AfectarAntesDe FROM configuracion WHERE Codigo = 1", CnnPrincipal, adOpenDynamic, adLockOptimistic
  If Val(rstUniversal!FechaAfectada) = 1 Then
    Dim Fecha As Date
    Fecha = DevFechaHoraServidor
    If Val(Format(Fecha, "H")) < Val(rstUniversal!AfectarAntesDe) Then
      DevFechaEspecial = Fecha - (0.042 * Val(rstUniversal!HorasAfectacion))
    Else
      DevFechaEspecial = Fecha
    End If
  Else
    DevFechaEspecial = DevFechaHoraServidor
  End If
  CerrarRecorset rstUniversal
End Function

Public Function DevConsecutivoGuiasFactura() As Boolean
  Dim rstConfiguracion As New ADODB.Recordset
  rstConfiguracion.CursorLocation = adUseClient
  AbrirRecorset rstConfiguracion, "SELECT ConsecutivoGuiasFactura FROM configuracion", CnnPrincipal, adOpenDynamic, adLockOptimistic
  If Val(rstConfiguracion.Fields("ConsecutivoGuiasFactura")) = 1 Then
    DevConsecutivoGuiasFactura = True
  Else
    DevConsecutivoGuiasFactura = False
  End If
  CerrarRecorset rstConfiguracion
  
End Function

Public Function DevImprimirGuiaFormato() As Boolean
  Dim rstConfiguracion As New ADODB.Recordset
  rstConfiguracion.CursorLocation = adUseClient
  AbrirRecorset rstConfiguracion, "SELECT ImprimirGuiaFormato FROM configuracion", CnnPrincipal, adOpenDynamic, adLockOptimistic
  If Val(rstConfiguracion.Fields("ImprimirGuiaFormato")) = 1 Then
    DevImprimirGuiaFormato = True
  Else
    DevImprimirGuiaFormato = False
  End If
  CerrarRecorset rstConfiguracion
  Set rstConfiguracion = Nothing
End Function


