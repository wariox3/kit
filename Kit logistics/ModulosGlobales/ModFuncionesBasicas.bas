Attribute VB_Name = "ModFuncionesBasicas"
Option Explicit
Public Const HH_DISPLAY_TOPIC = &H0
Const HH_SET_WIN_TYPE = &H4
Const HH_GET_WIN_TYPE = &H5
Const HH_GET_WIN_HANDLE = &H6
Const HH_DISPLAY_TEXT_POPUP = &HE   ' Display string resource ID or text in a pop-up window.
Const HH_HELP_CONTEXT = &HF         ' Display mapped numeric value in dwData.
Const HH_TP_HELP_CONTEXTMENU = &H10 ' Text pop-up help, similar to WinHelp's HELP_CONTEXTMENU.
Const HH_TP_HELP_WM_HELP = &H11     ' text pop-up help, similar to WinHelp's HELP_WM_HELP.
Const HH_CLOSE_ALL = &H12
Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long

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
Public Function ConvertirALetras(Numero) As String
  If Numero <> "" Then
  Dim I As Integer, L As String
  ConvertirALetras = ""
  For I = 1 To Len(Numero)
    L = Mid(Numero, I, 1)
    Select Case L
      Case "1"
        ConvertirALetras = ConvertirALetras & "A"
      Case "2"
        ConvertirALetras = ConvertirALetras & "B"
      Case "3"
        ConvertirALetras = ConvertirALetras & "C"
      Case "4"
        ConvertirALetras = ConvertirALetras & "D"
      Case "5"
        ConvertirALetras = ConvertirALetras & "E"
      Case "6"
        ConvertirALetras = ConvertirALetras & "F"
      Case "7"
        ConvertirALetras = ConvertirALetras & "G"
      Case "8"
        ConvertirALetras = ConvertirALetras & "H"
      Case "9"
        ConvertirALetras = ConvertirALetras & "I"
      Case "0"
        ConvertirALetras = ConvertirALetras & "J"
    End Select
  Next
  End If
End Function
