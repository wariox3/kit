VERSION 5.00
Begin VB.Form FrmAcercaDe 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acerca de Recogidas [Kit Logitics]"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   $"FrmAcercaDe.frx":0000
      Height          =   855
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "FrmAcercaDe.frx":0090
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "FrmAcercaDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdSalir_Click()
  Unload Me
End Sub
