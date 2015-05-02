VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmDevFechas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingrese las fechas..."
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DPFecha1 
      Height          =   300
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   23527427
      CurrentDate     =   38510
   End
   Begin MSComCtl2.DTPicker DPFecha2 
      Height          =   300
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   23527427
      CurrentDate     =   38510
   End
   Begin VB.Label LblMensaje 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label LblTitFecha2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha 2:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   630
   End
   Begin VB.Label LblTitFecha1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha 1:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   630
   End
End
Attribute VB_Name = "FrmDevFechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdAceptar_Click()
  LasFechas(1) = DPFecha1.Value
  LasFechas(2) = DPFecha2.Value
  Ok = True
  Unload Me
End Sub
Private Sub CmdCancelar_Click()
  Ok = False
  Unload Me
End Sub
Private Sub DPFecha1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub
Private Sub DPFecha2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub
Private Sub Form_Load()
  DPFecha1.Value = Date
  DPFecha2.Value = Date
  Ok = True
End Sub
