Attribute VB_Name = "ModFuncionesPropias"
Option Explicit



Sub MsgTit(Mensaje As String)
  Principal.LblMensaje = Mensaje
  Principal.TmPrincipal.Enabled = True
End Sub

Public Function DvRutaEsp(Ruta As String) As String
  Dim RutaSalida As String
  If Ruta <> "" Then
    For II = 1 To Len(Ruta)
      If Mid(Ruta, II, 1) = "\" Then
        RutaSalida = RutaSalida & "\\"
      Else
        RutaSalida = RutaSalida & Mid(Ruta, II, 1)
      End If
    Next
    DvRutaEsp = RutaSalida
  Else
    DvRutaEsp = ""
  End If
End Function
