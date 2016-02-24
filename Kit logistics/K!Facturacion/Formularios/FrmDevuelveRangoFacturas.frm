VERSION 5.00
Begin VB.Form FrmDevuelveRangoFacturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rango facturas"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CboTipo 
      Height          =   315
      ItemData        =   "FrmDevuelveRangoFacturas.frx":0000
      Left            =   1080
      List            =   "FrmDevuelveRangoFacturas.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox TxtHasta 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox TxtDesde 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   510
   End
End
Attribute VB_Name = "FrmDevuelveRangoFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
  If Val(TxtDesde) > 0 Then
    If Val(TxtHasta) > 0 Then
      If Val(TxtDesde) <= Val(TxtHasta) Then
        If CboTipo.ListIndex > -1 Then
          FufuLo = 1
          NumeroFacturaDesde = Val(TxtDesde)
          NumeroFacturaHasta = Val(TxtHasta)
          TipoFactura = (CboTipo.ListIndex) + 1
          Unload Me
        Else
          MsgBox "Debe escoger un tipo"
        End If
      Else
        MsgBox "El numero desde debe ser mayor al numero hasta"
      End If
    Else
      MsgBox "Debe especificar el numero hasta"
    End If
  Else
    MsgBox "Debe especificar el numero desde"
  End If
End Sub


Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  FufuLo = 0
End Sub
