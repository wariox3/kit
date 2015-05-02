Attribute VB_Name = "ModVariables"
Option Explicit
'*******************************************
Public CnnPrincipal As ADODB.Connection
Public rstUniversal As ADODB.Recordset
Public rstUniversalAux As ADODB.Recordset
'*******************************************
Public RutaLocal As String
Public ArchivoInf As String
Public ArchivoAyuda As String
Public Coperaciones As Long

Public CodUsuarioActivo As Integer
Public UsuarioActivo As String


Public II As Integer
Public Item As ListItem
Public FufuSt As String
Public FufuLo As Long
Public CooImp() As CoordenadasImpresion

Public Type CoordenadasImpresion
  VX As Integer
  VY As Integer
  Mostrar As Byte
  Tamaño As Byte
  Longitud As Integer
End Type

