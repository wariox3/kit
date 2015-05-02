Attribute VB_Name = "ModFunPro"
Option Explicit

Sub EnfocarT(CajaTexto As TextBox)
  CajaTexto.SelStart = 0
  CajaTexto.SelLength = Len(CajaTexto)
End Sub

