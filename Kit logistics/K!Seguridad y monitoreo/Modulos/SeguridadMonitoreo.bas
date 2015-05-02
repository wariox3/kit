Attribute VB_Name = "SeguridadMonitoreo"
Option Explicit

Public Function DevTipo(Tipo As Byte) As String
  Select Case Tipo
    Case 1
      DevTipo = "VIAJE"
    Case 2
      DevTipo = "REPARTO"
    Case 5
      DevTipo = "RECOGIDA"
  End Select
End Function
