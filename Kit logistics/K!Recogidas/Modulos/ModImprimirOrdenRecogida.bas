Attribute VB_Name = "ModImprimirOrdenRecogida"
Option Explicit
Sub ImprimirOrdenRecogida(Anuncio As Long)
Dim Inc As Long
IniciarImpresion (SacarDatoString(1))
  AbrirRecorset rstUniversal, "Select*from anuncios where IdAnuncio =" & Anuncio, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    UbC CooImp(1).Mostrar, CooImp(1).VX, CooImp(1).VY, CooImp(1).Tamaño, rstUniversal!IdAnuncio & "", CooImp(1).Longitud
    UbC CooImp(2).Mostrar, CooImp(2).VX, CooImp(2).VY, CooImp(2).Tamaño, rstUniversal!IdCliente & "", CooImp(2).Longitud
    UbC CooImp(4).Mostrar, CooImp(4).VX, CooImp(4).VY, CooImp(4).Tamaño, rstUniversal!TelAnunciante & "", CooImp(4).Longitud
    UbC CooImp(5).Mostrar, CooImp(5).VX, CooImp(5).VY, CooImp(5).Tamaño, rstUniversal!DirAnunciante & "", CooImp(5).Longitud
    UbC CooImp(6).Mostrar, CooImp(6).VX, CooImp(6).VY, CooImp(6).Tamaño, rstUniversal!FhRecogida & "", CooImp(6).Longitud
    UbC CooImp(7).Mostrar, CooImp(7).VX, CooImp(7).VY, CooImp(7).Tamaño, Format(rstUniversal!FhRecogida, "hh:mm") & "", CooImp(7).Longitud
    UbC CooImp(8).Mostrar, CooImp(8).VX, CooImp(8).VY, CooImp(8).Tamaño, rstUniversal!Anunciante & "", CooImp(8).Longitud
    UbC CooImp(9).Mostrar, CooImp(9).VX, CooImp(9).VY, CooImp(9).Tamaño, rstUniversal!Unidades & "", CooImp(9).Longitud
    UbC CooImp(10).Mostrar, CooImp(10).VX, CooImp(10).VY, CooImp(10).Tamaño, rstUniversal!KilosReales & "", CooImp(10).Longitud
    UbC CooImp(11).Mostrar, CooImp(11).VX, CooImp(11).VY, CooImp(11).Tamaño, rstUniversal!KilosVol & "", CooImp(11).Longitud
    UbC CooImp(12).Mostrar, CooImp(12).VX, CooImp(12).VY, CooImp(12).Tamaño, rstUniversal!IdRuta & "", CooImp(12).Longitud
    UbC CooImp(13).Mostrar, CooImp(13).VX, CooImp(13).VY, CooImp(13).Tamaño, UsuarioActivo, CooImp(13).Longitud
    FufuLo = SacarDatoString(2)
    FufuSt = rstUniversal!Comentarios & ""
    Inc = 0
    II = 1
    Do While II < Len(FufuSt)
      If Mid(FufuSt, II, 1) = " " Then ' para evitar espacios en blanco al comienzo
        II = II + 1
      Else
        If FufuLo <= 0 Then Exit Do
        FufuLo = FufuLo - 1
        UbC CooImp(14).Mostrar, CooImp(14).VX, CooImp(14).VY + Inc, CooImp(14).Tamaño, Mid(FufuSt, II, CooImp(14).Longitud), CooImp(14).Longitud
        II = II + CooImp(14).Longitud
        Inc = Inc + 4
      End If
    Loop
    FufuSt = rstUniversal!IdCliente & ""
    CerrarRecorset rstUniversal
    If FufuSt <> "0" Then
      AbrirRecorset rstUniversal, "Select NmCliente from Clientes where IdCliente ='" & FufuSt & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        UbC CooImp(3).Mostrar, CooImp(3).VX, CooImp(3).VY, CooImp(3).Tamaño, rstUniversal!NmCliente & "", CooImp(3).Longitud
      End If
      CerrarRecorset rstUniversal
    End If
Printer.EndDoc
TerminarImpresion
End Sub
