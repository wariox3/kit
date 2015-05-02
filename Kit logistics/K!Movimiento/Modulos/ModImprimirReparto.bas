Attribute VB_Name = "ModImprimirReparto"
Option Explicit

Sub ImprimirReparto(IdDespacho As Long)
  Dim Tipo As Byte
  IniciarImpresion (19)
  II = 0
  FufuLo = 1
  AbrirRecorset rstUniversal, "Select OrdDespacho, FhExpedicion, IdEncargado, Tipo from DespachosReparto where OrdDespacho=" & IdDespacho, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      Tipo = rstUniversal!Tipo
      FufuSt = rstUniversal!IdEncargado & ""
    End If
  CerrarRecorset rstUniversal
  AbrirRecorset rstUniversal, "Select Id, Nombre, Apellido1, Apellido2 from DatosBasicos where Id=" & FufuSt, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then FufuSt = rstUniversal!Nombre & " " & rstUniversal!Apellido1 & " " & rstUniversal!Apellido2
  CerrarRecorset rstUniversal
  AbrirRecorset rstUniversal, "SELECT Guias.Guia, Guias.Remitente, Guias.DocCliente, Guias.NmDestinatario, Ciudades.NmCiudad, Guias.Unidades, Guias.KilosReales, Guias.DespachoRep, Guias.Estado FROM Guias INNER JOIN Ciudades ON Guias.IdCiuDestino = Ciudades.IdCiudad Where (Guias.DespachoRep = " & IdDespacho & ") ORDER BY Guias.Guia", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  ImprimirEncabezado IdDespacho, Tipo
  Do While rstUniversal.EOF = False
    If FufuLo = 6 Then
      II = 0
      FufuLo = 1
      Printer.NewPage
      ImprimirEncabezado IdDespacho, Tipo
    End If
    UbC CooImp(5).Mostrar, CooImp(5).VX, CooImp(5).VY + II, CooImp(5).Tamaño, rstUniversal!Guia & "", CooImp(5).Longitud
    UbC CooImp(6).Mostrar, CooImp(6).VX, CooImp(6).VY + II, CooImp(6).Tamaño, rstUniversal!Remitente & "", CooImp(6).Longitud
    UbC CooImp(7).Mostrar, CooImp(7).VX, CooImp(7).VY + II, CooImp(7).Tamaño, rstUniversal!DocCliente & "", CooImp(7).Longitud
    UbC CooImp(8).Mostrar, CooImp(8).VX, CooImp(8).VY + II, CooImp(8).Tamaño, rstUniversal!NmDestinatario & "", CooImp(8).Longitud
    UbC CooImp(9).Mostrar, CooImp(9).VX, CooImp(9).VY + II, CooImp(9).Tamaño, rstUniversal!NmCiudad & "", CooImp(9).Longitud
    UbC CooImp(10).Mostrar, CooImp(10).VX, CooImp(10).VY + II, CooImp(10).Tamaño, rstUniversal!Unidades & "", CooImp(10).Longitud
    UbC CooImp(11).Mostrar, CooImp(11).VX, CooImp(11).VY + II, CooImp(11).Tamaño, rstUniversal!KilosReales & "", CooImp(11).Longitud
    II = II + 5
    FufuLo = FufuLo + 1
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
  Printer.EndDoc
  TerminarImpresion
  MsgBox "Se van a actualizar el estado de las guias", vbInformation
  AbrirRecorset rstUniversal, "Update Guias set Estado='V' where DespachoRep=" & IdDespacho, CnnPrincipal, adOpenDynamic, adLockOptimistic
End Sub
Sub ImprimirEncabezado(Despacho As Long, Tip As Byte)
  UbC CooImp(1).Mostrar, CooImp(1).VX, CooImp(1).VY, CooImp(1).Tamaño, Str(Despacho), CooImp(1).Longitud
  UbC CooImp(2).Mostrar, CooImp(2).VX, CooImp(2).VY, CooImp(2).Tamaño, Date, CooImp(2).Longitud
  UbC CooImp(3).Mostrar, CooImp(3).VX, CooImp(3).VY, CooImp(3).Tamaño, FufuSt, CooImp(3).Longitud
  UbC CooImp(4).Mostrar, CooImp(4).VX, CooImp(4).VY, CooImp(4).Tamaño, DevTipoDespacho(Tip), CooImp(4).Longitud
End Sub

Private Function DevTipoDespacho(ElTipo As Byte) As String
  Select Case ElTipo
    Case 0
      DevTipoDespacho = "REPARTO"
    Case 1
      DevTipoDespacho = "REEXPEDICION"
    Case 2
      DevTipoDespacho = "AUXILIAR"
  End Select
End Function
