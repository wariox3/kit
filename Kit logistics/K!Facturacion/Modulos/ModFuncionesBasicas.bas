Attribute VB_Name = "ModFuncionesBasicas"
Option Explicit
Sub EnfocarT(CajaTexto As TextBox)
  CajaTexto.SelStart = 0
  CajaTexto.SelLength = Len(CajaTexto)
End Sub
Public Sub EnfocarM(Caja As MaskEdBox)
  Caja.SelStart = 0
  Caja.SelLength = Len(Caja)
End Sub
Public Function ValNum(Caja As Variant) As Single
  If IsNumeric(Caja.Text) = False Then
    ValNum = 0
  Else
    ValNum = Caja
  End If
End Function
Sub IconosTool(Tool As Toolbar, ListaImagenes As ImageList)
  Tool.ImageList = ListaImagenes
  For II = 1 To 5
      Tool.Buttons(II + 2).Image = II
  Next
    Tool.Buttons(11).Image = 8
    Tool.Buttons(12).Image = 7
    Tool.Buttons(13).Image = 9
    Tool.Buttons(14).Image = 6
    Tool.Buttons(9).Image = 11
    Tool.Buttons(16).Image = 10
    Tool.Buttons(17).Image = 12
    Tool.Buttons(18).Image = 13
    Tool.Buttons(19).Image = 14
End Sub

Sub BotTool(Inicio As Byte, Fin As Byte, ElTool As Toolbar, Bloqueo As Boolean)
Dim I As Byte
If Bloqueo = True Then
  For I = Inicio To Fin
    If I <> 7 Then
      ElTool.Buttons.Item(I).Enabled = False
    End If
  Next I
  ElTool.Buttons.Item(4).Enabled = True
  ElTool.Buttons.Item(7).Enabled = True
  ElTool.Buttons.Item(18).Enabled = False
  ElTool.Buttons.Item(19).Enabled = False
Else
  For I = Inicio To Fin
    If I <> 7 Then
      ElTool.Buttons.Item(I).Enabled = True
    End If
  Next I
  ElTool.Buttons.Item(4).Enabled = False
  ElTool.Buttons.Item(7).Enabled = False
  ElTool.Buttons.Item(18).Enabled = True
  ElTool.Buttons.Item(19).Enabled = True
End If
End Sub
Public Sub LaTecla(LaTec As KeyCodeConstants, ElTool As Toolbar)
  Select Case LaTec
    Case vbKeyF3 ' Eliminar
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (6)
    Case vbKeyF4 ' Cancelar
      If ElTool.Buttons.Item(4).Enabled = True Then Screen.ActiveForm.AccionTool (7)
    Case vbKeyF5 ' Primero
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (11)
    Case vbKeyF6 ' Anterior
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (12)
    Case vbKeyF7 ' Siguiente
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (13)
    Case vbKeyF8 ' Ultimo
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (14)
    Case vbKeyF9 ' Nuevo
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (3)
    Case vbKeyF11 ' Guardar
      If ElTool.Buttons.Item(4).Enabled = True Then Screen.ActiveForm.AccionTool (4)
    Case vbKeyF10 ' Editar
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (5)
    Case vbKeyF12 ' Cerrar
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (16)
    Case vbKeyHome
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (9)
    Case vbKeyEnd ' Imprimir
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (18)
    Case vbKeyPause  ' Cargar informacion adicional
      If ElTool.Buttons.Item(4).Enabled = False Then Screen.ActiveForm.AccionTool (19)
  End Select
End Sub

