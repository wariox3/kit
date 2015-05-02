Attribute VB_Name = "FunPro"
Option Explicit

Sub IconosTool(Tool As Toolbar, ListaImagenes As ImageList)
  Tool.ImageList = ListaImagenes
  For II = 1 To 5
      Tool.Buttons(II + 2).Image = II
  Next
    Tool.Buttons(11).Image = 7
    Tool.Buttons(12).Image = 8
    Tool.Buttons(13).Image = 9
    Tool.Buttons(14).Image = 10
    Tool.Buttons(9).Image = 6
    Tool.Buttons(16).Image = 11
    Tool.Buttons(17).Image = 13
    Tool.Buttons(18).Image = 12
    Tool.Buttons(19).Image = 14
    'Tool.Buttons(20).Image = 15
End Sub

Sub EnfocarT(CajaTexto As TextBox)
  CajaTexto.SelStart = 0
  CajaTexto.SelLength = Len(CajaTexto)
End Sub
