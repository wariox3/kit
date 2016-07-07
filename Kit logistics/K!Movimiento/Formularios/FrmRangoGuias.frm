VERSION 5.00
Begin VB.Form FrmRangoGuias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rango guias"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox TxtHasta 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox TxtDesde 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   510
   End
End
Attribute VB_Name = "FrmRangoGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CmdCancelar_Click()
  FufuLo = 0
  Unload Me
End Sub

Private Sub Command1_Click()
  If Val(TxtDesde) <> 0 And Val(TxtHasta) <> 0 Then
    If Val(TxtDesde) < Val(TxtHasta) Then
      FufuLo = 1
      GuiaDesde = Val(TxtDesde)
      GuiaHasta = Val(TxtHasta)
      Unload Me
    Else
      MsgBox "Desde debe ser menor a hasta"
    End If
  Else
    MsgBox "El rango debe ser diferente de cero"
  End If
  
End Sub
