VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmOpciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Opciones"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5106
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Inicio"
      TabPicture(0)   =   "FrmOpciones.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Programar recogidas"
      TabPicture(1)   =   "FrmOpciones.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Otros"
      TabPicture(2)   =   "FrmOpciones.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame1 
         Caption         =   "Opciones"
         Height          =   1455
         Left            =   -74760
         TabIndex        =   4
         Top             =   600
         Width           =   4455
         Begin VB.CheckBox Check1 
            Caption         =   "Mostrar mensaje cuando se actualiza el listado de recogidas pendientes. (Recomendado)"
            Height          =   495
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   4095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Mostrar mensaje cuando se actualizan los vehiculos en transito de recogida. (Recomendado)"
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   840
            Width           =   3975
         End
      End
      Begin VB.Frame FraDatos 
         Height          =   855
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   2295
         Begin VB.CheckBox ChkRecPendiente 
            Caption         =   "Mostrar recogidas pendientes al iniciar"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "FrmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdAceptar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  ChkRecPendiente.Value = GetSetting("Kit Logistics", "Recogidas", "Ini_Rec_Pend", 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "Kit logistics", "Recogidas", "Ini_Rec_Pend", ChkRecPendiente.Value
End Sub
