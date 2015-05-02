Attribute VB_Name = "ModImprmirGuia"
Option Explicit
Dim rstImpGuia As New ADODB.Recordset

Public Function ImprimirGuia(Guia As Long) As Boolean
  Dim PagoDestino As Double, Inc As Integer, J As Integer, I As Integer
  rstImpGuia.CursorLocation = adUseClient
  If IniciarImpresion(13) = True Then
      AbrirRecorset rstImpGuia, "Select guias.*, usuarios.NmUsuario from guias LEFT JOIN usuarios on guias.IdUsuario = usuarios.IDUsuario where guia =" & Guia, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        
        If Val(rstImpGuia!GuiFac) = 1 Then
          UbC CooImp(34).Mostrar, CooImp(34).VX, CooImp(34).VY, CooImp(34).Tamaño, "A" & rstImpGuia!Guia & "", CooImp(34).Longitud
        Else
          UbC CooImp(1).Mostrar, CooImp(1).VX, CooImp(1).VY, CooImp(1).Tamaño, rstImpGuia!Guia & "", CooImp(1).Longitud
        End If
        UbC CooImp(15).Mostrar, CooImp(15).VX, CooImp(15).VY, CooImp(15).Tamaño, Format(rstImpGuia!FhEntradaBodega & "", "dd/mm/yy"), CooImp(15).Longitud
        UbC CooImp(2).Mostrar, CooImp(2).VX, CooImp(2).VY, CooImp(2).Tamaño, rstImpGuia!CR & "", CooImp(2).Longitud
        UbC CooImp(3).Mostrar, CooImp(3).VX, CooImp(3).VY, CooImp(3).Tamaño, rstImpGuia!Cliente & "", CooImp(3).Longitud
        'Impresion del numero de factura

        
        
        AbrirRecorset rstUniversal, "Select terceros.*, ciudades.NmCiudad from terceros, ciudades where (terceros.idciudad=ciudades.idciudad) and IdTercero='" & rstImpGuia!Cuenta & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
        If rstUniversal.EOF = False Then
          UbC CooImp(4).Mostrar, CooImp(4).VX, CooImp(4).VY, CooImp(4).Tamaño, rstUniversal!NmCiudad & "", CooImp(4).Longitud
          UbC CooImp(5).Mostrar, CooImp(5).VX, CooImp(5).VY, CooImp(5).Tamaño, rstUniversal!Direccion & "", CooImp(5).Longitud
          UbC CooImp(6).Mostrar, CooImp(6).VX, CooImp(6).VY, CooImp(6).Tamaño, rstUniversal!Telefono & "", CooImp(6).Longitud
        End If
        CerrarRecorset rstUniversal
        
        UbC CooImp(7).Mostrar, CooImp(7).VX, CooImp(7).VY, CooImp(7).Tamaño, rstImpGuia!Remitente & "", CooImp(7).Longitud
        UbC CooImp(10).Mostrar, CooImp(10).VX, CooImp(10).VY, CooImp(10).Tamaño, rstImpGuia!DocCliente & "", CooImp(10).Longitud
        UbC CooImp(11).Mostrar, CooImp(11).VX, CooImp(11).VY, CooImp(11).Tamaño, rstImpGuia!NmDestinatario & "", CooImp(11).Longitud
        UbC CooImp(14).Mostrar, CooImp(14).VX, CooImp(14).VY, CooImp(14).Tamaño, DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Val(rstImpGuia!IdCiuDestino & ""), "NmCiudad", CnnPrincipal), CooImp(14).Longitud
        UbC CooImp(12).Mostrar, CooImp(12).VX, CooImp(12).VY, CooImp(12).Tamaño, rstImpGuia!DirDestinatario & "", CooImp(12).Longitud
        UbC CooImp(13).Mostrar, CooImp(13).VX, CooImp(13).VY, CooImp(13).Tamaño, rstImpGuia!TelDestinatario & "", CooImp(13).Longitud
        FufuSt = DevResBus("SELECT IdTipoCobro, NmTipoCobro FROM tipos_cobro WHERE IdTipoCobro=" & Val(rstImpGuia!TipoCobro & ""), "NmTipoCobro", CnnPrincipal)
        UbC CooImp(26).Mostrar, CooImp(26).VX, CooImp(26).VY, CooImp(26).Tamaño, FufuSt, CooImp(26).Longitud
        UbC CooImp(27).Mostrar, CooImp(27).VX, CooImp(27).VY, CooImp(27).Tamaño, FufuSt, CooImp(27).Longitud
        
        PagoDestino = 0
        If (Val(rstImpGuia!TipoCobro) = 1 Or Val(rstImpGuia!TipoCobro) = 2) And Val(rstImpGuia!Recaudo) = 0 Then
          UbC CooImp(17).Mostrar, CooImp(17).VX, CooImp(17).VY, CooImp(17).Tamaño, rstImpGuia!VrFlete & "", CooImp(17).Longitud
          UbC CooImp(18).Mostrar, CooImp(18).VX, CooImp(18).VY, CooImp(18).Tamaño, rstImpGuia!VrManejo & "", CooImp(18).Longitud
          PagoDestino = PagoDestino + Val(rstImpGuia.Fields("VrFlete")) + Val(rstImpGuia.Fields("VrManejo"))
        Else
          UbC CooImp(17).Mostrar, CooImp(17).VX, CooImp(17).VY, CooImp(17).Tamaño, ConvertirALetras(rstImpGuia!VrFlete), CooImp(17).Longitud
          UbC CooImp(18).Mostrar, CooImp(18).VX, CooImp(18).VY, CooImp(18).Tamaño, ConvertirALetras(rstImpGuia!VrManejo), CooImp(18).Longitud
        End If
         
        If Val(rstImpGuia!TipoCobro) <> 1 And Val(rstImpGuia!TipoCobro) <> 2 And Val(rstImpGuia.Fields("Recaudo")) <> 0 Then
          UbC CooImp(32).Mostrar, CooImp(32).VX, CooImp(32).VY, CooImp(32).Tamaño, rstImpGuia!Recaudo & "", CooImp(32).Longitud
        Else
          UbC CooImp(32).Mostrar, CooImp(32).VX, CooImp(32).VY, CooImp(32).Tamaño, (PagoDestino - Val(rstImpGuia!Abonos)), CooImp(32).Longitud
        End If
        
        UbC CooImp(16).Mostrar, CooImp(16).VX, CooImp(16).VY, CooImp(16).Tamaño, ConvertirALetras(rstImpGuia!VrDeclarado), CooImp(16).Longitud
        
        FufuSt = ""
        FufuSt = FufuSt & rstImpGuia!Observaciones & ""
        Inc = 0
        I = 1
        J = 1
        Do While I < Len(FufuSt)
          If J = 7 Then Exit Do
          J = J + 1
          UbC CooImp(23).Mostrar, CooImp(23).VX, CooImp(23).VY + Inc, CooImp(23).Tamaño, Mid(FufuSt, I, CooImp(23).Longitud), CooImp(23).Longitud
          I = I + CooImp(23).Longitud
          Inc = Inc + 4
        Loop
  
        UbC CooImp(29).Mostrar, CooImp(29).VX, CooImp(29).VY, CooImp(29).Tamaño, rstImpGuia!Unidades & "", CooImp(29).Longitud
        UbC CooImp(30).Mostrar, CooImp(30).VX, CooImp(30).VY, CooImp(30).Tamaño, rstImpGuia!KilosReales & "", CooImp(30).Longitud
        UbC CooImp(31).Mostrar, CooImp(31).VX, CooImp(31).VY, CooImp(31).Tamaño, rstImpGuia!KilosVolumen & "", CooImp(31).Longitud
        UbC CooImp(33).Mostrar, CooImp(33).VX, CooImp(33).VY, CooImp(33).Tamaño, rstImpGuia!NmUsuario & "", CooImp(33).Longitud
        

        
      CerrarRecorset rstImpGuia
    Printer.EndDoc
    TerminarImpresion
    ImprimirGuia = True
  Else
    ImprimirGuia = False
  End If
End Function
