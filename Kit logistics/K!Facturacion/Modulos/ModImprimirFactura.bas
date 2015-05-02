Attribute VB_Name = "ModImprimirFactura"
Option Explicit
Dim I As Integer, J As Integer, Inc As Integer
Sub GImprimirFactura(IdFactura As Long, Tipo As Byte)
If IniciarImpresion(3) = True Then
  Printer.PaperSize = 1
  Dim Kilos As Long
  Dim Unidades As Long
  Dim Flete As Currency
  Dim Manejo As Currency
  Dim Otros As Currency
  Dim TipoFactura As Integer
  
  AbrirRecorset rstUniversal, "SELECT facturas.*, Terceros.* from Facturas, terceros where (facturas.idCliente=terceros.IdTercero) and IdFactura=" & IdFactura, CnnPrincipal, adOpenDynamic, adLockOptimistic
    TipoFactura = rstUniversal.Fields("IdTipoFactura")
    UbC CooImp(1).Mostrar, CooImp(1).VX, CooImp(1).VY, CooImp(1).Tamaño, Format(rstUniversal.Fields("FhFac"), "dd"), 2
    UbC CooImp(2).Mostrar, CooImp(2).VX, CooImp(2).VY, CooImp(2).Tamaño, Format(rstUniversal.Fields("FhFac"), "mm"), 2
    UbC CooImp(3).Mostrar, CooImp(3).VX, CooImp(3).VY, CooImp(3).Tamaño, Format(rstUniversal.Fields("FhFac"), "yy"), 2
    
    UbC CooImp(4).Mostrar, CooImp(4).VX, CooImp(4).VY, CooImp(4).Tamaño, Format(rstUniversal.Fields("FhVenceFac"), "dd"), 2
    UbC CooImp(5).Mostrar, CooImp(5).VX, CooImp(5).VY, CooImp(5).Tamaño, Format(rstUniversal.Fields("FhVenceFac"), "mm"), 5
    UbC CooImp(6).Mostrar, CooImp(6).VX, CooImp(6).VY, CooImp(6).Tamaño, Format(rstUniversal.Fields("FhVenceFac"), "yy"), 2
    
    UbC CooImp(29).Mostrar, CooImp(29).VX, CooImp(29).VY, CooImp(29).Tamaño, Str(IdFactura) & "-" & TipoFactura, CooImp(29).Longitud
    UbC CooImp(7).Mostrar, CooImp(7).VX, CooImp(7).VY, CooImp(7).Tamaño, rstUniversal.Fields("RazonSocial"), CooImp(7).Longitud
    UbC CooImp(8).Mostrar, CooImp(8).VX, CooImp(8).VY, CooImp(8).Tamaño, DigitoNIT(rstUniversal.Fields("IdTercero") & ""), CooImp(8).Longitud
    UbC CooImp(9).Mostrar, CooImp(9).VX, CooImp(9).VY, CooImp(9).Tamaño, rstUniversal.Fields("Direccion"), CooImp(9).Longitud
    UbC CooImp(10).Mostrar, CooImp(10).VX, CooImp(10).VY, CooImp(10).Tamaño, rstUniversal.Fields("Telefono"), CooImp(10).Longitud
    UbC CooImp(11).Mostrar, CooImp(11).VX, CooImp(11).VY, CooImp(11).Tamaño, DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Val(rstUniversal!IdCiudad), "NmCiudad", CnnPrincipal), CooImp(11).Longitud
    Flete = rstUniversal!TFlete
    Manejo = rstUniversal!TManejo
    Otros = rstUniversal!TOtros
    Inc = 0
    I = 1
    J = 1
    Do While I < Len(rstUniversal.Fields("Notas"))
      If J = 4 Then Exit Do
      J = J + 1
      UbC CooImp(23).Mostrar, CooImp(23).VX, CooImp(23).VY + Inc, CooImp(23).Tamaño, Mid(rstUniversal.Fields("Notas"), I, CooImp(23).Longitud), CooImp(23).Longitud
      I = I + CooImp(23).Longitud
      Inc = Inc + 4
    Loop
    
  CerrarRecorset rstUniversal
  
  Select Case Tipo
    Case 1
      UbC CooImp(40).Mostrar, CooImp(40).VX, CooImp(40).VY, CooImp(40).Tamaño, "GUIA    DOC CLTE     DESTINATA    DESTINO       UND  KILOS   DECLARA    MANEJO     FLETE", 90
      Select Case TipoFactura
        Case 1
          AbrirRecorset rstUniversal, "SELECT guias.*, ciudades.NmCiudad From Guias, ciudades where (guias.idciudestino=ciudades.idciudad) and IdFactura=" & IdFactura, CnnPrincipal, adOpenDynamic, adLockBatchOptimistic
        Case 2
          AbrirRecorset rstUniversal, "SELECT guias.*, ciudades.NmCiudad From Guias, ciudades where (guias.idciudestino=ciudades.idciudad) and IdFactura2=" & IdFactura, CnnPrincipal, adOpenDynamic, adLockBatchOptimistic
        Case 3
          AbrirRecorset rstUniversal, "SELECT guias.*, ciudades.NmCiudad From Guias, ciudades where (guias.idciudestino=ciudades.idciudad) and IdFactura3=" & IdFactura, CnnPrincipal, adOpenDynamic, adLockBatchOptimistic
      End Select

        I = 1
        Inc = 0
        Do While rstUniversal.EOF = False
          UbC CooImp(13).Mostrar, CooImp(13).VX, CooImp(13).VY + Inc, CooImp(13).Tamaño, rstUniversal.Fields("Guia"), CooImp(13).Longitud
          UbC CooImp(15).Mostrar, CooImp(15).VX, CooImp(15).VY + Inc, CooImp(15).Tamaño, rstUniversal.Fields("DocCliente") & "", CooImp(15).Longitud
          UbC CooImp(16).Mostrar, CooImp(16).VX, CooImp(16).VY + Inc, CooImp(16).Tamaño, rstUniversal.Fields("NmDestinatario"), CooImp(16).Longitud
          UbC CooImp(17).Mostrar, CooImp(17).VX, CooImp(17).VY + Inc, CooImp(17).Tamaño, rstUniversal.Fields("NmCiudad"), CooImp(17).Longitud
          UbC CooImp(46).Mostrar, CooImp(46).VX, CooImp(46).VY + Inc, CooImp(46).Tamaño, rstUniversal.Fields("EmpaqueRef") & "", CooImp(46).Longitud
          
          UbCN CooImp(18).Mostrar, CooImp(18).VX, CooImp(18).VY + Inc, CooImp(18).Tamaño, rstUniversal.Fields("Unidades"), CooImp(18).Longitud, "@@@@@"
          UbCN CooImp(19).Mostrar, CooImp(19).VX, CooImp(19).VY + Inc, CooImp(19).Tamaño, rstUniversal.Fields("KilosFacturados"), CooImp(19).Longitud, "@@@@@"
          UbCN CooImp(20).Mostrar, CooImp(20).VX, CooImp(20).VY + Inc, CooImp(20).Tamaño, rstUniversal.Fields("VrDeclarado"), CooImp(20).Longitud, "@@@@@@@@@@"
          UbCN CooImp(21).Mostrar, CooImp(21).VX, CooImp(21).VY + Inc, CooImp(21).Tamaño, Format(rstUniversal.Fields("VrFlete"), "#,##0;(#,##0)"), CooImp(21).Longitud, "@@@@@@@@@"
          UbCN CooImp(22).Mostrar, CooImp(22).VX, CooImp(22).VY + Inc, CooImp(22).Tamaño, Format(rstUniversal.Fields("VrManejo"), "#,##0;(#,##0)"), CooImp(22).Longitud, "@@@@@@@@@"

          Kilos = Kilos + rstUniversal.Fields("KilosFacturados")
          Unidades = Unidades + rstUniversal.Fields("Unidades")
          rstUniversal.MoveNext
          I = I + 1
          Inc = Inc + 4
        Loop
      CerrarRecorset rstUniversal
      
    Case 2
      UbC CooImp(40).Mostrar, CooImp(40).VX, CooImp(40).VY, CooImp(40).Tamaño, "PLANILLA    RELACION      FLETE        MANEJO       TOTAL", 90
      AbrirRecorset rstUniversal, "SELECT * from FacturasPlanillas where IdFactura=" & IdFactura, CnnPrincipal, adOpenDynamic, adLockBatchOptimistic
        I = 1
        Inc = 0
        Do While rstUniversal.EOF = False
          UbC CooImp(36).Mostrar, CooImp(36).VX, CooImp(36).VY + Inc, CooImp(36).Tamaño, rstUniversal.Fields("IdPlanilla"), CooImp(36).Longitud
          UbC CooImp(37).Mostrar, CooImp(37).VX, CooImp(37).VY + Inc, CooImp(37).Tamaño, rstUniversal.Fields("RelCliente") & "", CooImp(37).Longitud
          UbCN CooImp(38).Mostrar, CooImp(38).VX, CooImp(38).VY + Inc, CooImp(38).Tamaño, Format(rstUniversal.Fields("VrFletePlanilla"), "#,##0;(#,##0)"), CooImp(38).Longitud, "@@@@@@@@@"
          UbCN CooImp(39).Mostrar, CooImp(39).VX, CooImp(39).VY + Inc, CooImp(39).Tamaño, Format(rstUniversal.Fields("VrManejoPlanilla"), "#,##0;(#,##0)"), CooImp(39).Longitud, "@@@@@@@@@"
          UbCN CooImp(42).Mostrar, CooImp(42).VX, CooImp(42).VY + Inc, CooImp(42).Tamaño, Format(rstUniversal.Fields("VrManejoPlanilla") + rstUniversal.Fields("VrFletePlanilla"), "#,##0;(#,##0)"), CooImp(39).Longitud, "@@@@@@@@@"
          rstUniversal.MoveNext
          I = I + 1
          Inc = Inc + 4
        Loop
      CerrarRecorset rstUniversal
  
    Case 3
      UbC CooImp(40).Mostrar, CooImp(40).VX, CooImp(40).VY, CooImp(40).Tamaño, "CONCEPTO                                        VALOR", 90
      AbrirRecorset rstUniversal, "Select ConceptosFacturas.*, conceptoscontables.NmConcepto from ConceptosFacturas, ConceptosContables Where (ConceptosFacturas.IdConcepto=ConceptosContables.IdConcepto) and IdFactura=" & IdFactura, CnnPrincipal, adOpenDynamic, adLockBatchOptimistic
        I = 1
        Inc = 0
        Do While rstUniversal.EOF = False
          UbC CooImp(34).Mostrar, CooImp(34).VX, CooImp(34).VY + Inc, CooImp(34).Tamaño, rstUniversal.Fields("NmConcepto") & "", CooImp(34).Longitud
          UbCN CooImp(35).Mostrar, CooImp(35).VX, CooImp(35).VY + Inc, CooImp(35).Tamaño, Format(rstUniversal.Fields("Valor"), "#,##0;(#,##0)"), CooImp(35).Longitud, "@@@@@@@@@@"
            rstUniversal.MoveNext
          I = I + 1
          Inc = Inc + 4
        Loop
      CerrarRecorset rstUniversal
      
  
  End Select
  
  'AbrirRecorset rstUniversal, "SELECT FacturacionConceptos.IdFactura, Conceptos.NmConcepto, FacturacionConceptos.Valor FROM FacturacionConceptos LEFT JOIN Conceptos ON FacturacionConceptos.IdConcepto = Conceptos.IdConcepto where IdFactura=" & IdFactura, 2
  '  Do While rstUniversal.EOF = False
  
  '    rstUniversal.MoveNext
  '    i = i + 1
  '    Inc = Inc + 4
  '  Loop
  'CerrarRecorset rstUniversal
  '&&&&&&&&&&&&&&&&&&&&&&&& Conceptos &&&&&&&&&&&&&&&&&&&&&&&
  
    UbC CooImp(43).Mostrar, CooImp(43).VX, CooImp(43).VY, CooImp(43).Tamaño, "Flete: ", CooImp(43).Longitud
    UbC CooImp(44).Mostrar, CooImp(44).VX, CooImp(44).VY, CooImp(44).Tamaño, "Manejo: ", CooImp(44).Longitud
    UbC CooImp(45).Mostrar, CooImp(45).VX, CooImp(45).VY, CooImp(45).Tamaño, "Otros: ", CooImp(45).Longitud
    UbC CooImp(41).Mostrar, CooImp(41).VX, CooImp(41).VY, CooImp(41).Tamaño, "NOTAS ADICIONALES", 60
    
    UbCN CooImp(24).Mostrar, CooImp(24).VX, CooImp(24).VY, CooImp(24).Tamaño, Str(Unidades), CooImp(24).Longitud, "@@@@@@"
    UbCN CooImp(25).Mostrar, CooImp(25).VX, CooImp(25).VY, CooImp(25).Tamaño, Str(Kilos), CooImp(25).Longitud, "@@@@@@"
    UbCN CooImp(27).Mostrar, CooImp(27).VX, CooImp(27).VY, CooImp(27).Tamaño, Format(Flete, "#,##0;(#,##0)"), CooImp(27).Longitud, "@@@@@@@@@"
    UbCN CooImp(26).Mostrar, CooImp(26).VX, CooImp(26).VY, CooImp(26).Tamaño, Format(Manejo, "#,##0;(#,##0)"), CooImp(26).Longitud, "@@@@@@@@@"
    UbCN CooImp(31).Mostrar, CooImp(31).VX, CooImp(31).VY, CooImp(31).Tamaño, Format(Otros, "#,##0;(#,##0)"), CooImp(31).Longitud, "@@@@@@@@@"
    If Tipo = 3 Then
      UbCN CooImp(47).Mostrar, CooImp(47).VX, CooImp(47).VY, CooImp(47).Tamaño, Format(Otros, "#,##0;(#,##0)"), CooImp(47).Longitud, "@@@@@@@@@"
    End If
    UbCN CooImp(33).Mostrar, CooImp(33).VX, CooImp(33).VY, CooImp(33).Tamaño, Format(Flete + Manejo + Otros, "#,##0;(#,##0)"), CooImp(33).Longitud, "@@@@@@@@@"
    UbC CooImp(32).Mostrar, CooImp(32).VX, CooImp(32).VY, CooImp(32).Tamaño, UCase(CovLetras(Flete + Manejo + Otros)), CooImp(32).Longitud
  
  Printer.EndDoc
  Erase CooImp

End If
End Sub


