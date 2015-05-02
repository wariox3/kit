VERSION 5.00
Begin VB.Form FrmDevDirectorio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Devuelve directorio"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   3690
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   7215
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   4320
      Width           =   1935
   End
End
Attribute VB_Name = "FrmDevDirectorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
  FufuSt = Dir1.Path
  Unload Me
End Sub

Private Sub CmdCancelar_Click()
  FufuSt = ""
  Unload Me
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub
