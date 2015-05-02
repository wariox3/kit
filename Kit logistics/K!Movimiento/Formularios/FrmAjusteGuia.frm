VERSION 5.00
Begin VB.Form FrmAjusteGuia 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ajuste de guia"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Frame FraDatos 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.TextBox TxtComentarios 
         Height          =   1605
         Left            =   600
         TabIndex        =   8
         Top             =   1560
         Width           =   4095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Text            =   "Text3"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Declarado:"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   5
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Comentarios:"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   4
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Manejo:"
         Height          =   195
         Index           =   1
         Left            =   375
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Flete:"
         Height          =   195
         Index           =   0
         Left            =   555
         TabIndex        =   2
         Top             =   240
         Width           =   915
      End
   End
End
Attribute VB_Name = "FrmAjusteGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

