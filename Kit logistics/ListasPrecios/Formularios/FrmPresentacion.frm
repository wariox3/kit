VERSION 5.00
Begin VB.Form FrmPresentacion 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmPresentacion.frx":0000
   ScaleHeight     =   4755
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1200
      Left            =   240
      Top             =   120
   End
   Begin VB.Label LblPropietario 
      BackStyle       =   0  'Transparent
      Caption         =   "Propietario"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label LblEmpresa 
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   2880
      Width           =   3615
   End
   Begin VB.Label LblIdProducto 
      BackStyle       =   0  'Transparent
      Caption         =   "Producto"
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   3120
      Width           =   3615
   End
End
Attribute VB_Name = "FrmPresentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  LblEmpresa = GetSetting("Kit Logistics", "InfoSoftware", "Empresa", "Sin empresa")
  LblPropietario = GetSetting("Kit Logistics", "InfoSoftware", "Propietario", "Sin Propietario")
  LblIdProducto = GetSetting("Kit Logistics", "InfoSoftware", "Serial", "Sin serial")
    
End Sub

Private Sub Timer1_Timer()
  Unload Me
  Principal.Show
End Sub
