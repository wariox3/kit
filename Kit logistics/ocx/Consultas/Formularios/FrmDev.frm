VERSION 5.00
Begin VB.Form FrmDev 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingrese la informacion...."
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdReglas 
      Caption         =   "Reglas"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox TxtIngreso 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label LblMensaje 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
End
Attribute VB_Name = "FrmDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
  Ok = False
  Select Case Tip
    Case 1
      If Len(TxtIngreso.Text) <= 5 Then
        MsgBox "Las contraseñas deben tener mas de 5 digitos"
        Ok = False
      Else
        ElSt = TxtIngreso.Text
        Ok = True
      End If
    Case 2
      If TxtIngreso.Text = "" Then
        MsgBox "El numero de caracteres no puede ser 0"
        Ok = False
      Else
        ElSt = TxtIngreso.Text
        Ok = True
      End If
    Case 3
      If Val(TxtIngreso) > 2147483648# Then
        MsgBox "El valor para este campo debe se menor a 2.147.483.648"
        TxtIngreso.Text = ""
        Ok = False
      Else
        If Val(TxtIngreso.Text) = 0 Then
          MsgBox "El valor no puede se cero (0)", vbCritical
        Else
          Ello = Val(TxtIngreso)
          Ok = True
        End If
      End If
    Case 4
      If Len(TxtIngreso.Text) <> 6 Then
        MsgBox "La placa debe ser de 6 digitos"
        Ok = False
      Else
        ElSt = UCase(TxtIngreso.Text)
        Ok = True
      End If
    Case 5
      If Len(TxtIngreso.Text) > 10 Then
        MsgBox "El maximo de digitos para el nit es de 10"
        Ok = False
      Else
        ElSt = UCase(TxtIngreso.Text)
        Ok = True
      End If
      
    Case 6
      TxtIngreso.Text = "1" & TxtIngreso.Text
      If Val(TxtIngreso) > 2147483648# Then
        MsgBox "El valor para este campo debe se menor a 2.147.483.648"
        TxtIngreso.Text = ""
        Ok = False
      Else
        If Val(TxtIngreso.Text) = 0 Then
          MsgBox "El valor no puede se cero (0)", vbCritical
        Else
          Ello = Val(TxtIngreso)
          Ok = True
        End If
      End If
  End Select
  If Ok = True Then
    Unload Me
  End If
End Sub
Private Sub CmdCancelar_Click()
  Ok = False
  Unload Me
End Sub


Private Sub TxtIngreso_GotFocus()
  TxtIngreso.SelStart = 0
  TxtIngreso.SelLength = Len(TxtIngreso)
End Sub

Private Sub TxtIngreso_KeyPress(KeyAscii As Integer)
  Select Case Tip
    Case 4
      If KeyAscii = 32 Then
        MsgBox "No se permiten espacios"
        KeyAscii = 0
      End If
    Case 5
      If InStr("0123456789" + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0: MsgBox "Solo se permiten numeros en este campo"
  End Select
End Sub
