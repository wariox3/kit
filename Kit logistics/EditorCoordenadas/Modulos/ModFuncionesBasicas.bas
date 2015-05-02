Attribute VB_Name = "ModFuncionesBasicas"
Option Explicit
 Type TpCoordenadasImpresion
      Campo As String
      VX As Integer
      VY As Integer
      Mostrar As Byte
      Tamaño As Byte
      Longitud As Integer
      Descripcion As String
End Type
Public Function CpExisteFichero(Ruta As String) As Boolean
  Dim x
  On Error GoTo ErrorHandler:
  x = GetAttr(Ruta)
  CpExisteFichero = True
  Exit Function
ErrorHandler:
  CpExisteFichero = False
End Function

