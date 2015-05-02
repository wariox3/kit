Attribute VB_Name = "ModFuncionesImpresion"
Option Explicit

Public Function IniciarImpresion(Tipo As Integer) As Boolean
Dim NR As Byte, Ori As Byte, Alto As Currency, Ancho As Currency
Dim strRutaCoordenadasImpresion As String
  Select Case Tipo
      'Facturas
    Case 3
      strRutaCoordenadasImpresion = GetSetting("Kit Logistics", "Facturacion", "CoordenadasImpresionFactura")
    'Guias
    Case 13
      strRutaCoordenadasImpresion = GetSetting("Kit Logistics", "Movimiento", "CoordenadasImpresionGuia")
    'Recibo de caja
    Case 16
      strRutaCoordenadasImpresion = GetSetting("Kit Logistics", "Movimiento", "CoordenadasImpresionReciboCaja")
    'Planilla Reparto
    Case 19
      strRutaCoordenadasImpresion = GetSetting("Kit Logistics", "Movimiento", "CoordenadasImpresionPlanillaReparto")
    'Manifiesto
    Case 15
      strRutaCoordenadasImpresion = GetSetting("Kit Logistics", "Movimiento", "CoordenadasImprresionManifiesto")
  End Select
  If CpExisteFichero(strRutaCoordenadasImpresion) = True Then
    Open strRutaCoordenadasImpresion For Input As #1
      Input #1, FufuSt, NR, Ori, Alto, Ancho
      ReDim CooImp(NR)
      For II = 1 To NR
        Input #1, FufuSt, CooImp(II).VX, CooImp(II).VY, CooImp(II).Mostrar, CooImp(II).Tamaño, FufuSt, CooImp(II).Longitud
      Next
      Close #1
      Printer.ScaleMode = 6
      'Printer.Height = Alto
      'Printer.Width = Ancho
      'Printer.ScaleLeft = 1
      'Printer.ScaleTop = 1
      'Printer.Orientation = Ori
      Printer.FontName = "Courier New"
      IniciarImpresion = True
  Else
    MsgBox "El archivo de coordenadas para imprimir " & strRutaCoordenadasImpresion & " no existe... Comuniquese con el proveedor del software": Exit Function
    IniciarImpresion = False
  End If
End Function
Sub TerminarImpresion()
   Erase CooImp
End Sub
Public Sub UbC(Mostrar As Byte, x As Integer, y As Integer, TmLetra As Byte, Texto As String, Longitud As Integer)
  If Mostrar = 1 Then
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.FontSize = TmLetra
    Printer.Print Mid(Texto, 1, Longitud)
  End If
End Sub

Public Sub UbCN(Mostrar As Byte, x As Integer, y As Integer, TmLetra As Byte, Texto As String, Longitud As Integer, NroCaracteres As String)
  If Mostrar = 1 Then
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.FontSize = TmLetra
    Printer.Print Format(Mid(Texto, 1, Longitud), NroCaracteres)
  End If
End Sub
