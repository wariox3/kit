Attribute VB_Name = "ModImprimirManifiesto"
Option Explicit
Dim ManDestino As String, ManOrigen As String
Sub ImprimirManifiesto(IdDespacho As Long)
Dim Datos(2) As String, Observaciones As String, Inc As Integer, J As Integer, I As Integer
If IniciarImpresion(15) = True Then
  establecerPapel
  MsgTit "Sacando datos del manifiesto"
  AbrirRecorset rstUniversal, "Select*from Despachos where OrdDespacho=" & IdDespacho, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Datos(0) = rstUniversal!IdVehiculo & ""
    Datos(1) = rstUniversal!IdConductor & ""
    UbC CooImp(1).Mostrar, CooImp(1).VX, CooImp(1).VY, CooImp(1).Tamaño, "[M-" & rstUniversal!IdManifiesto & "/D-" & IdDespacho & "]", CooImp(1).Longitud
    UbC CooImp(1).Mostrar, CooImp(1).VX, CooImp(1).VY + 6, CooImp(1).Tamaño, "[ME-" & rstUniversal!ManElectronico & "]", 50
    ManOrigen = DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Val(rstUniversal!IdCiudadOrigen), "NmCiudad", CnnPrincipal)
    ManDestino = DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Val(rstUniversal!IdCiudadDestino), "NmCiudad", CnnPrincipal)
    UbC CooImp(2).Mostrar, CooImp(2).VX, CooImp(2).VY, CooImp(2).Tamaño, Format(Date, "dd/mm/yy"), CooImp(2).Longitud
    UbC CooImp(3).Mostrar, CooImp(3).VX, CooImp(3).VY, CooImp(3).Tamaño, ManOrigen, CooImp(3).Longitud
    UbC CooImp(4).Mostrar, CooImp(4).VX, CooImp(4).VY, CooImp(4).Tamaño, ManDestino, CooImp(4).Longitud
    
    UbCN CooImp(46).Mostrar, CooImp(46).VX, CooImp(46).VY, CooImp(46).Tamaño, Format(rstUniversal!VrFlete, "#,##0;(#,##0)"), CooImp(46).Longitud, "@@@@@@@@@@"
    UbCN CooImp(47).Mostrar, CooImp(47).VX, CooImp(47).VY, CooImp(47).Tamaño, Format(rstUniversal!VrDctoRteFte, "#,##0;(#,##0)"), CooImp(47).Longitud, "@@@@@@@@@@"
    UbCN CooImp(48).Mostrar, CooImp(48).VX, CooImp(48).VY, CooImp(48).Tamaño, Format(rstUniversal!VrDctoIndCom, "#,##0;(#,##0)"), CooImp(48).Longitud, "@@@@@@@@@@"
    UbCN CooImp(49).Mostrar, CooImp(49).VX, CooImp(49).VY, CooImp(49).Tamaño, Format(Val(rstUniversal!VrFlete) - (Val(rstUniversal!VrDctoRteFte) + Val(rstUniversal!VrDctoIndCom)), "#,##0;(#,##0)"), CooImp(49).Longitud, "@@@@@@@@@@"
    UbCN CooImp(50).Mostrar, CooImp(50).VX, CooImp(50).VY, CooImp(50).Tamaño, Format(rstUniversal!VrAnticipo, "#,##0;(#,##0)"), CooImp(50).Longitud, "@@@@@@@@@@"
    UbCN CooImp(51).Mostrar, CooImp(51).VX, CooImp(51).VY, CooImp(51).Tamaño, Format(Val(rstUniversal!VrFlete) - (Val(rstUniversal!VrDctoRteFte) + Val(rstUniversal!VrDctoIndCom) + Val(rstUniversal!VrAnticipo)), "#,##0;(#,##0)"), CooImp(51).Longitud, "@@@@@@@@@@"
    UbC CooImp(52).Mostrar, CooImp(52).VX, CooImp(52).VY, CooImp(52).Tamaño, CovLetras(rstUniversal!VrFlete), CooImp(52).Longitud
    
    MsgTit "Posisionando coordenadas..."
    Dim rstInformacionEmpresa As New ADODB.Recordset
    rstInformacionEmpresa.CursorLocation = adUseClient
    AbrirRecorset rstInformacionEmpresa, "SELECT informacionempresa.* FROM informacionempresa WHERE Id = 1", CnnPrincipal, adOpenDynamic, adLockOptimistic
    UbC CooImp(54).Mostrar, CooImp(54).VX, CooImp(54).VY, CooImp(54).Tamaño, rstInformacionEmpresa!CodRegionalMin & "", CooImp(54).Longitud
    UbC CooImp(55).Mostrar, CooImp(55).VX, CooImp(55).VY, CooImp(55).Tamaño, rstInformacionEmpresa!CodEmpresaMin & "", CooImp(55).Longitud
    UbC CooImp(56).Mostrar, CooImp(56).VX, CooImp(56).VY, CooImp(56).Tamaño, "", CooImp(56).Longitud
    UbC CooImp(57).Mostrar, CooImp(57).VX, CooImp(57).VY, CooImp(57).Tamaño, "", CooImp(57).Longitud
    UbC CooImp(58).Mostrar, CooImp(58).VX, CooImp(58).VY, CooImp(58).Tamaño, "", CooImp(58).Longitud
    UbC CooImp(59).Mostrar, CooImp(59).VX, CooImp(59).VY, CooImp(59).Tamaño, "", CooImp(59).Longitud
    UbC CooImp(60).Mostrar, CooImp(60).VX, CooImp(60).VY, CooImp(60).Tamaño, "", CooImp(60).Longitud
    UbC CooImp(61).Mostrar, CooImp(61).VX, CooImp(61).VY, CooImp(61).Tamaño, "", CooImp(61).Longitud
    UbC CooImp(62).Mostrar, CooImp(62).VX, CooImp(62).VY, CooImp(62).Tamaño, rstInformacionEmpresa!Aseguradora & "", CooImp(62).Longitud
    UbC CooImp(63).Mostrar, CooImp(63).VX, CooImp(63).VY, CooImp(63).Tamaño, rstInformacionEmpresa!NroPoliza & "", CooImp(63).Longitud
    UbC CooImp(64).Mostrar, CooImp(64).VX, CooImp(64).VY, CooImp(64).Tamaño, rstInformacionEmpresa!VencePoliza & "", CooImp(64).Longitud
    UbC CooImp(65).Mostrar, CooImp(65).VX, CooImp(65).VY, CooImp(65).Tamaño, "", CooImp(65).Longitud
    CerrarRecorset rstInformacionEmpresa
    Observaciones = "CE[" & Format(Val(rstUniversal.Fields("FleteCE")) + Val(rstUniversal.Fields("TRecaudo")), "#,##0;(#,##0") & "] UND[" & rstUniversal.Fields("Unidades") & "] GUIAS[" & rstUniversal.Fields("Remesas") & "] " & rstUniversal.Fields("Observaciones") & ""
    Inc = 0
    I = 1
    J = 1
    Do While I < Len(Observaciones)
      If J = 9 Then Exit Do
      J = J + 1
      UbC CooImp(53).Mostrar, CooImp(53).VX, CooImp(53).VY + Inc, CooImp(53).Tamaño, Mid(Observaciones, I, CooImp(53).Longitud), 36
      I = I + CooImp(53).Longitud
      Inc = Inc + 4
    Loop
    
  CerrarRecorset rstUniversal

  '******** Conductor *********
  MsgTit "Informacion del conductor..."
  AbrirRecorset rstUniversalAux, "SELECT IdConductor, CONCAT(Conductores.Nombre, ' ', Conductores.Apellido1, ' ', Conductores.Apellido2) as NombreCompleto, Barrio, Categoria FROM Conductores where IdConductor='" & Datos(1) & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversalAux.EOF = False Then
      UbC CooImp(30).Mostrar, CooImp(30).VX, CooImp(30).VY, CooImp(30).Tamaño, rstUniversalAux!NombreCompleto & "", CooImp(30).Longitud
      UbC CooImp(31).Mostrar, CooImp(31).VX, CooImp(31).VY, CooImp(31).Tamaño, rstUniversalAux!IdConductor & "", CooImp(31).Longitud
      UbC CooImp(32).Mostrar, CooImp(32).VX, CooImp(32).VY, CooImp(32).Tamaño, rstUniversalAux!Barrio & "", CooImp(32).Longitud
      UbC CooImp(33).Mostrar, CooImp(33).VX, CooImp(33).VY, CooImp(33).Tamaño, rstUniversalAux!Categoria & "", CooImp(33).Longitud
    End If
  CerrarRecorset rstUniversalAux
  
  '******** Vehiculo *********
  PosVehiculo (Datos(0))
  'Printer.EndDoc
  ImprimirRemesasManifiesto (IdDespacho)
  TerminarImpresion
End If
End Sub
Private Sub PosVehiculo(Placa As String)
  MsgTit "Procesando informacion del vehiculo..."
  AbrirRecorset rstUniversalAux, "Select IdMarca, IdLinea, Modelo, ModeloRep, Serie, IdColor, IdCarroceria, RegNalCarga, PesoVacio, Soat, IdAseguradora, VenceSoat, PlacaRemolque, IdPropietario, IdTenedor from Vehiculos where IdPlaca='" & Placa & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    UbC CooImp(5).Mostrar, CooImp(5).VX, CooImp(5).VY, CooImp(5).Tamaño, Placa, CooImp(5).Longitud
    UbC CooImp(6).Mostrar, CooImp(6).VX, CooImp(6).VY, CooImp(6).Tamaño, DevResBus("SELECT IdMarca, NmMarca From Marcas where IdMarca=" & Val(rstUniversalAux!IdMarca), "NmMarca", CnnPrincipal), CooImp(6).Longitud
    UbC CooImp(7).Mostrar, CooImp(7).VX, CooImp(7).VY, CooImp(7).Tamaño, DevResBus("SELECT IdLinea, NmLinea From Lineas where IdLinea=" & Val(rstUniversalAux!IdLinea), "NmLinea", CnnPrincipal), CooImp(7).Longitud
    UbC CooImp(8).Mostrar, CooImp(8).VX, CooImp(8).VY, CooImp(8).Tamaño, rstUniversalAux!Modelo & "", CooImp(8).Longitud
    UbC CooImp(9).Mostrar, CooImp(9).VX, CooImp(9).VY, CooImp(9).Tamaño, rstUniversalAux!ModeloRep & "", CooImp(9).Longitud
    UbC CooImp(10).Mostrar, CooImp(10).VX, CooImp(10).VY, CooImp(10).Tamaño, rstUniversalAux!Serie & "", CooImp(10).Longitud
    UbC CooImp(11).Mostrar, CooImp(11).VX, CooImp(11).VY, CooImp(11).Tamaño, DevResBus("SELECT IdColor, NmColor From Colores where IdColor=" & Val(rstUniversalAux!IdColor), "NmColor", CnnPrincipal), CooImp(11).Longitud
    UbC CooImp(12).Mostrar, CooImp(12).VX, CooImp(12).VY, CooImp(12).Tamaño, DevResBus("SELECT IdCarroceria, NmCarroceria From Carrocerias where Idcarroceria=" & Val(rstUniversalAux!IdCarroceria), "NmCarroceria", CnnPrincipal), CooImp(12).Longitud
    UbC CooImp(13).Mostrar, CooImp(13).VX, CooImp(13).VY, CooImp(13).Tamaño, rstUniversalAux!RegNalCarga & "", CooImp(13).Longitud
    UbC CooImp(15).Mostrar, CooImp(15).VX, CooImp(15).VY, CooImp(15).Tamaño, rstUniversalAux!PesoVacio & "", CooImp(15).Longitud
    UbC CooImp(16).Mostrar, CooImp(16).VX, CooImp(16).VY, CooImp(16).Tamaño, rstUniversalAux!Soat & "", CooImp(16).Longitud
    UbC CooImp(17).Mostrar, CooImp(17).VX, CooImp(17).VY, CooImp(17).Tamaño, DevNombreDatosBasicos(rstUniversalAux!IdAseguradora), CooImp(17).Longitud
    UbC CooImp(18).Mostrar, CooImp(18).VX, CooImp(18).VY, CooImp(18).Tamaño, rstUniversalAux!VenceSoat & "", CooImp(18).Longitud
    UbC CooImp(19).Mostrar, CooImp(19).VX, CooImp(19).VY, CooImp(19).Tamaño, rstUniversalAux!PlacaRemolque & "", CooImp(19).Longitud
      '********* Propietario ************
        MsgTit "Datos del propietario..."
        AbrirRecorset rstUniversal, "SELECT Terceros.*, Ciudades.NmCiudad FROM Terceros, Ciudades where (Terceros.IdCiudad=Ciudades.IdCiudad) and IdTercero='" & rstUniversalAux!IdPropietario & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
          If rstUniversal.EOF = False Then
            UbC CooImp(20).Mostrar, CooImp(20).VX, CooImp(20).VY, CooImp(20).Tamaño, rstUniversal!RazonSocial & "", CooImp(20).Longitud
            UbC CooImp(21).Mostrar, CooImp(21).VX, CooImp(21).VY, CooImp(21).Tamaño, rstUniversal!IdTercero & "", CooImp(21).Longitud
            UbC CooImp(22).Mostrar, CooImp(22).VX, CooImp(22).VY, CooImp(22).Tamaño, rstUniversal!Direccion & "", CooImp(22).Longitud
            UbC CooImp(23).Mostrar, CooImp(23).VX, CooImp(23).VY, CooImp(23).Tamaño, rstUniversal!Telefono & "", CooImp(23).Longitud
            UbC CooImp(24).Mostrar, CooImp(24).VX, CooImp(24).VY, CooImp(24).Tamaño, rstUniversal!NmCiudad & "", CooImp(24).Longitud
          End If
        CerrarRecorset rstUniversal

      '********* Tenedor ****************
        MsgTit "Datos del tenedor..."
        AbrirRecorset rstUniversal, "SELECT Terceros.*, Ciudades.NmCiudad FROM Terceros, Ciudades where (Terceros.IdCiudad=Ciudades.IdCiudad) and IdTercero='" & rstUniversalAux!IdTenedor & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
          If rstUniversal.EOF = False Then
            UbC CooImp(25).Mostrar, CooImp(25).VX, CooImp(25).VY, CooImp(25).Tamaño, rstUniversal!RazonSocial & "", CooImp(25).Longitud
            UbC CooImp(26).Mostrar, CooImp(26).VX, CooImp(26).VY, CooImp(26).Tamaño, rstUniversal!IdTercero & "", CooImp(26).Longitud
            UbC CooImp(27).Mostrar, CooImp(27).VX, CooImp(27).VY, CooImp(27).Tamaño, rstUniversal!Direccion & "", CooImp(27).Longitud
            UbC CooImp(28).Mostrar, CooImp(28).VX, CooImp(28).VY, CooImp(28).Tamaño, rstUniversal!Telefono & "", CooImp(28).Longitud
            UbC CooImp(29).Mostrar, CooImp(29).VX, CooImp(29).VY, CooImp(29).Tamaño, rstUniversal!NmCiudad & "", CooImp(29).Longitud
          End If
        CerrarRecorset rstUniversal
  CerrarRecorset rstUniversalAux
End Sub

Private Sub ImprimirRemesasManifiesto(IdDespacho As Long)
  Dim rstRegistros As New ADODB.Recordset, NroRegistros As Double, I As Integer, Inc As Integer
  rstRegistros.CursorLocation = adUseClient
  
  rstRegistros.Open "Select Guia, IdDespacho From guias where IdDespacho=" & IdDespacho, CnnPrincipal, adOpenDynamic, adLockOptimistic
    NroRegistros = rstRegistros.RecordCount
  rstRegistros.Close
  
  If NroRegistros <= 10 Then
    rstRegistros.Open "SELECT guias.*, ciudades.NmCiudad From guias, ciudades where (guias.IdCiuDestino=ciudades.IdCiudad) and IdDespacho=" & IdDespacho, CnnPrincipal, adOpenDynamic, adLockBatchOptimistic
    I = 1
    Inc = 0
    Do While rstRegistros.EOF = False
      If I = 11 Then Exit Do
      UbC CooImp(35).Mostrar, CooImp(35).VX, CooImp(35).VY + Inc + 50, CooImp(35).Tamaño, Val(rstRegistros.Fields("guia")), CooImp(35).Longitud
      UbC CooImp(36).Mostrar, CooImp(36).VX, CooImp(36).VY + Inc + 50, CooImp(36).Tamaño, "Kilos", CooImp(36).Longitud
      UbC CooImp(37).Mostrar, CooImp(37).VX, CooImp(37).VY + Inc + 50, CooImp(37).Tamaño, Val(rstRegistros.Fields("Unidades")), CooImp(37).Longitud
      UbC CooImp(38).Mostrar, CooImp(38).VX, CooImp(38).VY + Inc + 50, CooImp(38).Tamaño, rstRegistros.Fields("KilosReales"), CooImp(38).Longitud
      UbC CooImp(39).Mostrar, CooImp(39).VX, CooImp(39).VY + Inc + 50, CooImp(39).Tamaño, "1", CooImp(39).Longitud
      UbC CooImp(40).Mostrar, CooImp(40).VX, CooImp(40).VY + Inc + 50, CooImp(40).Tamaño, "1", CooImp(40).Longitud
      UbC CooImp(41).Mostrar, CooImp(41).VX, CooImp(41).VY + Inc + 50, CooImp(41).Tamaño, "9880", CooImp(41).Longitud
      UbC CooImp(42).Mostrar, CooImp(42).VX, CooImp(42).VY + Inc + 50, CooImp(42).Tamaño, "Varios", CooImp(42).Longitud
      UbC CooImp(43).Mostrar, CooImp(43).VX, CooImp(43).VY + Inc + 50, CooImp(43).Tamaño, rstRegistros.Fields("NmCiudad"), CooImp(43).Longitud
      UbC CooImp(44).Mostrar, CooImp(44).VX, CooImp(44).VY + Inc + 50, CooImp(44).Tamaño, rstRegistros.Fields("Cliente") & "", CooImp(44).Longitud
      UbC CooImp(45).Mostrar, CooImp(45).VX, CooImp(45).VY + Inc + 50, CooImp(45).Tamaño, rstRegistros.Fields("NmDestinatario") & "", CooImp(45).Longitud
      rstRegistros.MoveNext
      I = I + 1
      Inc = Inc + 5
    Loop
    rstRegistros.Close
  End If
  
  Printer.EndDoc
  If NroRegistros > 10 Then
    GImprimirManifiestoDetalle IdDespacho
  End If
End Sub
Sub GImprimirManifiestoDetalle(IdDespacho As Long)
  Dim rstGuiasManifiesto As New ADODB.Recordset, Inc As Integer, J As Byte, TotPag As Integer, I As Integer
  rstGuiasManifiesto.CursorLocation = adUseClient
  MsgTit "Clasificando remisiones para el detalle..."
  IniciarImpresion (15)
  Printer.PaperSize = 1
  rstGuiasManifiesto.Open "SELECT guias.*, ciudades.NmCiudad From guias, ciudades where (guias.IdCiuDestino=ciudades.IdCiudad) and IdDespacho=" & IdDespacho, CnnPrincipal, adOpenDynamic, adLockOptimistic
  Inc = 0
  TotPag = 1
  J = 32
  For I = 33 To rstGuiasManifiesto.RecordCount
    J = J + 1
    If J = 33 Then TotPag = TotPag + 1: J = 0
  Next I
  If rstGuiasManifiesto.RecordCount > 0 Then
      J = 1
      I = 0
      Inc = 3
      MsgBox "Inserte papel con formato del detalle de manifiesto para la pagina " & J
      UbC 1, 100, 25, 10, "Pagina " & J & " De " & TotPag, 30
      'UbC CooImp(1).Mostrar, CooImp(1).VX, CooImp(1).VY, CooImp(1).Tamaño, "[" & "D-" & IdDespacho & "]", CooImp(1).Longitud
      'UbC CooImp(65).Mostrar, CooImp(65).VX, CooImp(65).VY, CooImp(65).Tamaño, Str(IdDespacho), CooImp(65).Longitud
      'UbC CooImp(67).Mostrar, CooImp(67).VX, CooImp(67).VY, CooImp(67).Tamaño, Format(Date, "dd mm yyyy"), CooImp(67).Longitud
      'UbC CooImp(68).Mostrar, CooImp(68).VX, CooImp(68).VY, CooImp(68).Tamaño, ManOrigen, CooImp(68).Longitud
      'UbC CooImp(66).Mostrar, CooImp(66).VX, CooImp(66).VY, CooImp(66).Tamaño, ManDestino, CooImp(66).Longitud
      
      'UbC CooImp(54).Mostrar, CooImp(54).VX, CooImp(54).VY, CooImp(54).Tamaño, SacarDatos(20, ArchivoInf), CooImp(54).Longitud
      'UbC CooImp(55).Mostrar, CooImp(55).VX, CooImp(55).VY, CooImp(55).Tamaño, SacarDatos(21, ArchivoInf), CooImp(55).Longitud
      'UbC CooImp(56).Mostrar, CooImp(56).VX, CooImp(56).VY, CooImp(56).Tamaño, SacarDatos(22, ArchivoInf), CooImp(56).Longitud

    
      Do While rstGuiasManifiesto.EOF = False
        If I = 33 Then
          If rstGuiasManifiesto.RecordCount > 33 Then
            Printer.NewPage
            Inc = 3
            I = 1
            J = J + 1
            MsgBox "Inserte papel con formato del detalle de manifiesto para la pagina " & J
            UbC 1, 100, 25, 10, "Pagina " & J & " De " & TotPag, 30
            'UbC CooImp(67).Mostrar, CooImp(67).VX, CooImp(67).VY, CooImp(67).Tamaño, Format(Date, "dd mm yyyy"), CooImp(67).Longitud
            'UbC CooImp(68).Mostrar, CooImp(68).VX, CooImp(68).VY, CooImp(68).Tamaño, ManOrigen, CooImp(68).Longitud
            'UbC CooImp(66).Mostrar, CooImp(66).VX, CooImp(66).VY, CooImp(66).Tamaño, ManDestino, CooImp(66).Longitud
            
            'UbC CooImp(54).Mostrar, CooImp(54).VX, CooImp(54).VY, CooImp(54).Tamaño, SacarDatos(20, ArchivoInf), CooImp(54).Longitud
            'UbC CooImp(55).Mostrar, CooImp(55).VX, CooImp(55).VY, CooImp(55).Tamaño, SacarDatos(21, ArchivoInf), CooImp(55).Longitud
            'UbC CooImp(56).Mostrar, CooImp(56).VX, CooImp(56).VY, CooImp(56).Tamaño, SacarDatos(22, ArchivoInf), CooImp(56).Longitud
            
          
          Else
            Exit Do
          End If
        End If
        UbC CooImp(35).Mostrar, CooImp(35).VX, CooImp(35).VY + Inc, CooImp(35).Tamaño, rstGuiasManifiesto.Fields("Guia"), CooImp(35).Longitud
        UbC CooImp(36).Mostrar, CooImp(36).VX, CooImp(36).VY + Inc, CooImp(36).Tamaño, "Kilos", CooImp(36).Longitud
        UbC CooImp(37).Mostrar, CooImp(37).VX, CooImp(37).VY + Inc, CooImp(37).Tamaño, rstGuiasManifiesto.Fields("Unidades"), CooImp(37).Longitud
        UbC CooImp(38).Mostrar, CooImp(38).VX, CooImp(38).VY + Inc, CooImp(38).Tamaño, rstGuiasManifiesto.Fields("KilosReales"), CooImp(38).Longitud
        UbC CooImp(39).Mostrar, CooImp(39).VX, CooImp(39).VY + Inc, CooImp(39).Tamaño, "1", CooImp(39).Longitud
        UbC CooImp(40).Mostrar, CooImp(40).VX, CooImp(40).VY + Inc, CooImp(40).Tamaño, "1", CooImp(40).Longitud
        UbC CooImp(41).Mostrar, CooImp(41).VX, CooImp(41).VY + Inc, CooImp(41).Tamaño, "9880", CooImp(41).Longitud
        UbC CooImp(42).Mostrar, CooImp(42).VX, CooImp(42).VY + Inc, CooImp(42).Tamaño, "Varios", CooImp(42).Longitud
        UbC CooImp(43).Mostrar, CooImp(43).VX, CooImp(43).VY + Inc, CooImp(43).Tamaño, rstGuiasManifiesto.Fields("NmCiudad"), CooImp(43).Longitud
        UbC CooImp(44).Mostrar, CooImp(44).VX, CooImp(44).VY + Inc, CooImp(44).Tamaño, rstGuiasManifiesto.Fields("Cliente"), CooImp(44).Longitud
        UbC CooImp(45).Mostrar, CooImp(45).VX, CooImp(45).VY + Inc, CooImp(45).Tamaño, rstGuiasManifiesto.Fields("NmDestinatario"), CooImp(45).Longitud
        rstGuiasManifiesto.MoveNext
        I = I + 1
        Inc = Inc + 5
      Loop
  End If
  rstGuiasManifiesto.Close
  '************************* Remesas ************************
  Printer.EndDoc
  TerminarImpresion
End Sub
