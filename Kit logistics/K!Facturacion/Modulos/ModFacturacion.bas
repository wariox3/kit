Attribute VB_Name = "ModFacturacion"
Option Explicit

Public Function CpEstFactura(IdFactura As Long) As String
Dim rstCp As New ADODB.Recordset
rstCp.CursorLocation = adUseClient
rstCp.Open "Select IdFactura, Estado from facturas where IdFactura=" & IdFactura, CnnPrincipal, adOpenDynamic, adLockOptimistic
If rstCp.EOF = False Then
  CpEstFactura = rstCp.Fields("Estado")
End If
Set rstCp = Nothing
End Function
