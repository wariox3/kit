Attribute VB_Name = "ModFuncionesImpresion"
Option Explicit
Sub IniciarImpresion(ArchivoCoordenadas As String)
Dim NR As Byte, Ori As Byte, Alto As Currency, Ancho As Currency
  If CpExisteFichero(ArchivoCoordenadas) = True Then
    Open ArchivoCoordenadas For Input As #1
      Input #1, FufuSt, NR, Ori, Alto, Ancho
      ReDim CooImp(NR)
      For II = 1 To NR
        Input #1, FufuSt, CooImp(II).VX, CooImp(II).VY, CooImp(II).Mostrar, CooImp(II).Tamaño, FufuSt, CooImp(II).Longitud
      Next
      Close #1
      Printer.ScaleMode = 6
      Printer.Height = Alto
      Printer.Width = Ancho
      Printer.ScaleLeft = 1
      Printer.ScaleTop = 1
      Printer.Orientation = Ori
      Printer.FontName = "Courier New"
  Else
    MsgBox "El archivo de coordenadas para imprimir la remision no existe... Comuniquese con el proveedor del software": Exit Sub
  End If
End Sub
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
