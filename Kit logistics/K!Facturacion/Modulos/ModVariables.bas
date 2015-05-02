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
