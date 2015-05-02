VERSION 5.00
Begin VB.Form FrmAcercaDe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de Ke!software Kelp 1.0.0"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "FrmAcercaDe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdInfoSis 
      Caption         =   "Info. del sitema..."
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   6
      X1              =   120
      X2              =   6000
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label LblTitulos 
      Caption         =   $"FrmAcercaDe.frx":000C
      Height          =   1275
      Index           =   7
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   4125
   End
   Begin VB.Label LblTitulos 
      AutoSize        =   -1  'True
      Caption         =   "Copyright ©     2005-2006 Ke!software inc."
      Height          =   195
      Index           =   6
      Left            =   1560
      TabIndex        =   6
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label LblTitulos 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Index           =   5
      Left            =   1560
      TabIndex        =   5
      Top             =   1440
      Width           =   4080
   End
   Begin VB.Label LblTitulos 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Index           =   4
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Width           =   4080
   End
   Begin VB.Label LblTitulos 
      AutoSize        =   -1  'True
      Caption         =   "Para la edicion de formatos de datos con precios"
      Height          =   195
      Index           =   3
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   3450
   End
   Begin VB.Label LblTitulos 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Top             =   1920
      Width           =   4080
   End
   Begin VB.Label LblTitulos 
      AutoSize        =   -1  'True
      Caption         =   "Se autoriza el uso de este producto a:"
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   2685
   End
   Begin VB.Label LblTitulos 
      AutoSize        =   -1  'True
      Caption         =   "Ke!software Kelp 1.0"
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1470
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   7
      X1              =   120
      X2              =   6000
      Y1              =   2400
      Y2              =   2400
   End
End
Attribute VB_Name = "FrmAcercaDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  'LblTitulos(5) = SacarDatoString(2)
  'LblTitulos(4) = SacarDatoString(11)
  'LblTitulos(2) = SacarDatoString(12)
End Sub
