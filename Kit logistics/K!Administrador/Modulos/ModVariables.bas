Attribute VB_Name = "ModVariables"
Option Explicit
Public Item As ListItem
Public II As Long
Public FufuSt As String
Public FufuLo As Long
Public ArchivoInf As String
Public FormAbierto As Boolean

Public CnnPrincipal As New ADODB.Connection
Public rstUniversal As New ADODB.Recordset

