Attribute VB_Name = "ModVariables"
Option Explicit

'*******************************************
Public CnnPrincipal As ADODB.Connection
Public rstUniversal As ADODB.Recordset
'*******************************************

Public CodUsuarioActivo As Integer
Public UsuarioActivo As String
Public Coperaciones As Long
Public FufuSt As String
Public CooImp() As CoordenadasImpresion

Public II As Integer
Public Item As ListItem
Public FufuLo As Long

Public Type CoordenadasImpresion
  VX As Integer
  VY As Integer
  Mostrar As Byte
  Tamaño As Byte
  Longitud As Integer
End Type

Public Type parametrosCartera
  Generar As Boolean
  GenerarExcel As Boolean
  sql As String
  IdCliente As Long
  IdAsesor As Long
  Numero As Long
  Tipo As Integer
  IdCentroOperaciones As Integer
End Type

Public Type parametrosRecibo
  Generar As Boolean
  GenerarExcel As Boolean
  sql As String
  IdCliente As Long
  Numero As Long
  Tipo As Integer
  Fecha As Boolean
  FechaDesde As String
  FechaHasta As String
  InformeDetallado As Boolean
End Type

Public Type parametrosNotaCredito
  Generar As Boolean
  GenerarExcel As Boolean
  sql As String
  IdCliente As Long
  Numero As Long
  Tipo As Integer
  Fecha As Boolean
  FechaDesde As String
  FechaHasta As String
  InformeDetallado As Boolean
End Type

Public varParametrosCartera As parametrosCartera
Public varParametrosRecibo As parametrosRecibo
Public varParametrosNotaCredito As parametrosNotaCredito
Public NumeroFacturaDesde As Long
Public NumeroFacturaHasta As Long
Public TipoFactura As Integer
