Attribute VB_Name = "ModImprimirRecCaja"
Option Explicit
Dim ValorT As Currency
Dim i As Integer, J As Integer
Dim rstTemporal As New ADODB.Recordset
Dim rstRecibos As New ADODB.Recordset
Dim rstUsuarios As New ADODB.Recordset

Sub GImprimirRecCajaEntrada(Guia As Long)
Dim Inc As Integer, Pagar As Double
Dim mystream As ADODB.Stream
Set mystream = New ADODB.Stream
mystream.Type = adTypeBinary

rstTemporal.CursorLocation = adUseClient
IniciarImpresion 16
rstTemporal.Open "SELECT Guia, FhEntradaBodega, IdCiuDestino, Cliente, Cuenta, Unidades, VrDeclarado, DocCLiente, NmDestinatario, DirDestinatario, IdUsuario, IdTpCtaFlete, IdTpCtaManejo, Ciudades.NmCiudad, Nombre, Nit, Direccion, Telefono, Email, Logo " & _
                 "FROM guias " & _
                 "LEFT JOIN ciudades ON guias.IdCiudestino = ciudades.IdCiudad " & _
                 "LEFT JOIN informacionempresa ON guias.IdEmpresa = informacionempresa.id " & _
                 "WHERE Guia = " & Guia, CnnPrincipal, adOpenDynamic, adLockOptimistic
rstRecibos.Open "Select recibos.* from recibos where GuiaRecibo=" & Guia, CnnPrincipal, adOpenDynamic, adLockOptimistic
rstUsuarios.Open "Select usuarios.* from usuarios where IDUsuario=" & Val(rstTemporal.Fields("IdUsuario")), CnnPrincipal, adOpenDynamic, adLockOptimistic
UbC 1, 70, 10, 16, rstTemporal.Fields("Nombre") & "", 50
UbC 1, 45, 17, 11, rstTemporal.Fields("Direccion") & "   Tel: " & rstTemporal.Fields("Telefono") & "   Nit. " & rstTemporal.Fields("Nit") & "", 80
UbC 1, 60, 22, 11, "e-mail/correo: " & rstTemporal.Fields("Email") & "", 50
If CpExisteFichero("c:\\logoempresa.gif") <> True Then
  mystream.Open
  mystream.Write rstTemporal.Fields("Logo")
  mystream.SaveToFile "c:\logoempresa.gif", adSaveCreateOverWrite
  mystream.Close
End If

Principal.Image1.Picture = LoadPicture("c:\\logoempresa.gif")
Printer.PaintPicture Principal.Image1, 5, 5, 40, 20
Principal.Image1.Picture = LoadPicture("")

UbC 1, 15, 25, 11, "_______________________________________________________________________________", 100
UbC CooImp(1).Mostrar, CooImp(1).VX, CooImp(1).VY, CooImp(1).Tamaño, "Recibo " & rstRecibos.Fields("NroRecibo"), CooImp(1).Longitud
UbC CooImp(6).Mostrar, CooImp(6).VX, CooImp(6).VY, CooImp(6).Tamaño, "Remision:    " & rstRecibos.Fields("GuiaRecibo"), CooImp(6).Longitud

UbC CooImp(2).Mostrar, CooImp(2).VX, CooImp(2).VY, CooImp(2).Tamaño, "Fecha:        Medellin, " & Format(rstTemporal.Fields("FhEntradaBodega"), "dd mmm yyyy"), CooImp(2).Longitud
UbC CooImp(3).Mostrar, CooImp(3).VX, CooImp(3).VY, CooImp(3).Tamaño, "Cliente:      " & rstTemporal.Fields("Cliente") & "", CooImp(3).Longitud
UbC CooImp(4).Mostrar, CooImp(4).VX, CooImp(4).VY, CooImp(4).Tamaño, "Nit Cliente:  " & rstTemporal.Fields("Cuenta") & "", CooImp(4).Longitud
UbC CooImp(5).Mostrar, CooImp(5).VX, CooImp(5).VY, CooImp(5).Tamaño, "Doc Cliente: " & rstTemporal.Fields("DocCliente") & "", CooImp(5).Longitud
UbC CooImp(8).Mostrar, CooImp(8).VX, CooImp(8).VY, CooImp(8).Tamaño, "Destinatario: " & rstTemporal.Fields("NmDestinatario") & "", CooImp(8).Longitud
UbC CooImp(9).Mostrar, CooImp(9).VX, CooImp(9).VY, CooImp(9).Tamaño, "Direccion:    " & rstTemporal.Fields("DirDestinatario"), CooImp(9).Longitud
UbC CooImp(10).Mostrar, CooImp(10).VX, CooImp(10).VY, CooImp(10).Tamaño, "Destino:     " & rstTemporal.Fields("NmCiudad") & "", CooImp(10).Longitud

UbC CooImp(15).Mostrar, CooImp(15).VX, CooImp(15).VY, CooImp(15).Tamaño, "Declarado:   " & Format(rstTemporal.Fields("VrDeclarado"), "$#,##0;($#,##0)"), CooImp(15).Longitud
UbC CooImp(16).Mostrar, CooImp(16).VX, CooImp(16).VY, CooImp(16).Tamaño, "Unidades:    " & rstTemporal.Fields("Unidades"), CooImp(16).Longitud

rstTemporal.Close

UbC 1, 15, 78, 11, "_______________________________________________________________________________", 100

UbC 1, 15, 84, 10, "Imputacion  Concepto               Debito        Credito", 100
AbrirRecorset rstTemporal, "SELECT IdTpCtaFlete, IdTpCtaManejo, VrFlete, VrManejo, IdUsuario FROM Guias where Guia=" & Guia, CnnPrincipal, adOpenDynamic, adLockOptimistic
  Inc = 0
  Pagar = 0
  If Val(rstTemporal.Fields("IdTpCtaFlete")) = 1 Then
    UbC CooImp(11).Mostrar, CooImp(11).VX, CooImp(11).VY, CooImp(11).Tamaño, "41450510", CooImp(11).Longitud
    UbC CooImp(12).Mostrar, CooImp(12).VX, CooImp(12).VY, CooImp(12).Tamaño, "FLETE CONTADO CLIENTE", CooImp(12).Longitud
    UbC CooImp(14).Mostrar, CooImp(14).VX, CooImp(14).VY, CooImp(14).Tamaño, Format(rstRecibos.Fields("Flete"), "$#,##0;($#,##0)"), CooImp(14).Longitud
    Inc = Inc + 5
    Pagar = Pagar + Val(rstRecibos.Fields("Flete"))
  End If
  
  If Val(rstTemporal.Fields("IdTpCtaManejo")) = 1 Then
    UbC CooImp(11).Mostrar, CooImp(11).VX, CooImp(11).VY + Inc, CooImp(11).Tamaño, "41454005", CooImp(11).Longitud
    UbC CooImp(12).Mostrar, CooImp(12).VX, CooImp(12).VY + Inc, CooImp(12).Tamaño, "COSTO MANRJO DE MCIA", CooImp(12).Longitud
    UbC CooImp(14).Mostrar, CooImp(14).VX, CooImp(14).VY + Inc, CooImp(14).Tamaño, Format(rstRecibos.Fields("Manejo"), "$#,##0;($#,##0)"), CooImp(14).Longitud
    Inc = Inc + 5
    Pagar = Pagar + Val(rstRecibos.Fields("Manejo"))
  End If
  
  If Val(rstTemporal.Fields("IdTpCtaManejo")) = 1 Then
    UbC CooImp(11).Mostrar, CooImp(11).VX, CooImp(11).VY + Inc, CooImp(11).Tamaño, "11050505", CooImp(11).Longitud
    UbC CooImp(12).Mostrar, CooImp(12).VX, CooImp(12).VY + Inc, CooImp(12).Tamaño, "CAJA", CooImp(12).Longitud
    UbC CooImp(13).Mostrar, CooImp(13).VX, CooImp(13).VY + Inc, CooImp(13).Tamaño, Format(Pagar, "$#,##0;($#,##0)"), CooImp(13).Longitud
  End If
UbC CooImp(18).Mostrar, CooImp(18).VX, CooImp(18).VY, CooImp(18).Tamaño, CovLetras(Val(Pagar)), CooImp(18).Longitud
UbC CooImp(19).Mostrar, CooImp(19).VX, CooImp(19).VY, CooImp(19).Tamaño, "Elaborado por " & rstUsuarios.Fields("NmUsuario") & "", CooImp(19).Longitud

rstTemporal.Close
Set rstRecibos = Nothing
Set rstUsuarios = Nothing
Printer.EndDoc
End Sub

Sub GImprimirRecCajaEntrada1(IdRec As Long)
'Dim ValorT As Currency
'IniciarImpresion 4
'AbrirRecorset rstUniversal, "SELECT RecibosCaja.*, Cliente.*, Ciudades.NmCiudad, Remisiones.Unidades, Remisiones.DocCliente, Remisiones.IdEntradaBodega, Remisiones.VrDeclarado, Remisiones.NmDestinatario, Remisiones.DirDestinatario FROM ((RecibosCaja LEFT JOIN ConsultaBasDireccion AS Cliente ON RecibosCaja.IdCliente = Cliente.ID) LEFT JOIN Remisiones ON RecibosCaja.IdEntradaBodega = Remisiones.IdEntradaBodega) LEFT JOIN Ciudades ON Remisiones.IdCiudadDestino = Ciudades.IdCiudad where IdRecibo=" & IdRec, 2
'UbC 1, 40, 10, 16, "DISTRIBUCION Y TRANSPORTES CUARTAS S.A", 50
'UbC 1, 45, 17, 11, "CRA. 56A # 62-63   Tel: 211 92 47   Nit. 8110104165", 80
'UbC 1, 65, 22, 11, "e-mail transcuartasa@epm.net.co", 50
'UbC 1, 15, 25, 11, "_______________________________________________________________________________", 100
'UbC CooImp(1).Mostrar, CooImp(1).VX, CooImp(1).VY, CooImp(1).Tamaño, "Recibo # " & CpVal(rstUniversal.Fields("IdRecibo")), CooImp(1).Longitud
'UbC CooImp(2).Mostrar, CooImp(2).VX, CooImp(2).VY, CooImp(2).Tamaño, "Fecha:        Medellin, " & Format(rstUniversal.Fields("Fecha"), "dd mmm yyyy"), CooImp(2).Longitud
'UbC CooImp(3).Mostrar, CooImp(3).VX, CooImp(3).VY, CooImp(3).Tamaño, "Cliente:      " & CpVal(rstUniversal.Fields("Nombre")), CooImp(3).Longitud
'UbC CooImp(4).Mostrar, CooImp(4).VX, CooImp(4).VY, CooImp(4).Tamaño, "Nit Cliente:  " & CpVal(rstUniversal.Fields("IdCliente")), CooImp(4).Longitud
'UbC CooImp(5).Mostrar, CooImp(5).VX, CooImp(5).VY, CooImp(5).Tamaño, "Doc Cliente: " & CpVal(rstUniversal.Fields("DocCliente")), CooImp(5).Longitud
'UbC CooImp(6).Mostrar, CooImp(6).VX, CooImp(6).VY, CooImp(6).Tamaño, "Remision:       " & CpVal(rstUniversal.Fields("Remisiones.IdEntradaBodega")), CooImp(6).Longitud
'UbC CooImp(7).Mostrar, CooImp(7).VX, CooImp(7).VY, CooImp(7).Tamaño, "Direccion:    " & CpVal(rstUniversal.Fields("BasDireccion")), CooImp(7).Longitud
'UbC CooImp(8).Mostrar, CooImp(8).VX, CooImp(8).VY, CooImp(8).Tamaño, "Destinatario: " & CpVal(rstUniversal.Fields("NmDestinatario")), CooImp(8).Longitud
'UbC CooImp(9).Mostrar, CooImp(9).VX, CooImp(9).VY, CooImp(9).Tamaño, "Direccion:    " & CpVal(rstUniversal.Fields("DirDestinatario")), CooImp(9).Longitud
'UbC CooImp(10).Mostrar, CooImp(10).VX, CooImp(10).VY, CooImp(10).Tamaño, "Destino:     " & CpVal(rstUniversal.Fields("NmCiudad")), CooImp(10).Longitud
'
'UbC CooImp(15).Mostrar, CooImp(15).VX, CooImp(15).VY, CooImp(15).Tamaño, "Declarado:   " & Format(CpVal(rstUniversal.Fields("VrDeclarado")), "$#,##0;($#,##0)"), CooImp(15).Longitud
'UbC CooImp(16).Mostrar, CooImp(16).VX, CooImp(16).VY, CooImp(16).Tamaño, "Unidades:    " & CpVal(rstUniversal.Fields("Unidades")), CooImp(16).Longitud
'UbC CooImp(17).Mostrar, CooImp(17).VX, CooImp(17).VY, CooImp(17).Tamaño, "Valor:   " & Format(CpVal(rstUniversal.Fields("Valor")), "$#,##0;($#,##0)"), CooImp(17).Longitud
'UbC CooImp(18).Mostrar, CooImp(18).VX, CooImp(18).VY, CooImp(18).Tamaño, CovLetras(rstUniversal.Fields("Valor")), CooImp(18).Longitud
'CerrarRecorset rstUniversal
'UbC 1, 15, 78, 11, "_______________________________________________________________________________", 100
'
'UbC 1, 15, 84, 10, "Imputacion  Concepto               Debito        Credito", 100
'AbrirRecorset rstUniversal, "SELECT RecibosCajaConceptos.*, Conceptos.NmConcepto, Conceptos.IdTpConcepto, Conceptos.IdCuenta FROM RecibosCajaConceptos LEFT JOIN Conceptos ON RecibosCajaConceptos.IdConcepto = Conceptos.IdConcepto where IdRecCaja=" & IdRec, 2
'  Dim Inc As Integer
'  I = 1
'  Inc = 0
'  Do While rstUniversal.EOF = False
'    If I = 4 Then Exit Do
'    UbC CooImp(11).Mostrar, CooImp(11).VX, CooImp(11).VY + Inc, CooImp(11).Tamaño, CpVal(rstUniversal.Fields("IdCuenta")), CooImp(11).Longitud
'    UbC CooImp(12).Mostrar, CooImp(12).VX, CooImp(12).VY + Inc, CooImp(12).Tamaño, CpVal(rstUniversal.Fields("NmConcepto")), CooImp(12).Longitud
'      If rstUniversal.Fields("IdTpConcepto") = 1 Then
'        UbC CooImp(13).Mostrar, CooImp(13).VX, CooImp(13).VY + Inc, CooImp(13).Tamaño, Format(CpVal(rstUniversal.Fields("Valor")), "$#,##0;($#,##0)"), CooImp(13).Longitud
'      Else
'        UbC CooImp(14).Mostrar, CooImp(14).VX, CooImp(14).VY + Inc, CooImp(14).Tamaño, Format(CpVal(rstUniversal.Fields("Valor")), "$#,##0;($#,##0)"), CooImp(14).Longitud
'      End If
'    ValorT = Val(CpVal(rstUniversal.Fields("Valor")))
'    rstUniversal.MoveNext
'    I = I + 1
'    Inc = Inc + 5
'  Loop
'  AbrirRecorset rstUniversal1, "SELECT Parametrizacion.CtoCaja, Conceptos.NmConcepto, Conceptos.IdCuenta FROM Parametrizacion LEFT JOIN Conceptos ON Parametrizacion.CtoCaja = Conceptos.IdConcepto", 2
'    UbC CooImp(11).Mostrar, CooImp(11).VX, CooImp(11).VY + Inc, CooImp(11).Tamaño, CpVal(rstUniversal1.Fields("IdCuenta")), CooImp(11).Longitud
'    UbC CooImp(12).Mostrar, CooImp(12).VX, CooImp(12).VY + Inc, CooImp(12).Tamaño, CpVal(rstUniversal1.Fields("NmConcepto")), CooImp(12).Longitud
'    UbC CooImp(14).Mostrar, CooImp(14).VX, CooImp(14).VY + Inc, CooImp(14).Tamaño, Format(ValorT, "$#,##0;($#,##0)"), CooImp(14).Longitud
'  CerrarRecorset rstUniversal1
'
'CerrarRecorset rstUniversal
'UbC CooImp(19).Mostrar, CooImp(19).VX, CooImp(19).VY + Inc, CooImp(19).Tamaño, "Elaborado por " & UsuarioActivo, CooImp(19).Longitud
'Printer.EndDoc

End Sub

