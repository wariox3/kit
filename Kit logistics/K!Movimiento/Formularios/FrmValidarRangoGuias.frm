VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmValidarRangoGuias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Validar rango guias"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   3480
      Width           =   1575
   End
   Begin MSComctlLib.ListView LstGuias 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Guia"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label LblHasta 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label LblDesde 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   510
   End
End
Attribute VB_Name = "FrmValidarRangoGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSalir_Click()
  Unload Me
End Sub



Private Sub Form_Load()
  LblDesde = FufuLo
  LblHasta = FufuLo2
  Ver
End Sub

Private Sub Ver()
AbrirRecorset rstUniversal, "SELECT guias.Guia from guias where Guia >= " & Val(LblDesde) & " AND Guia <=" & Val(LblHasta), CnnPrincipal, adOpenForwardOnly, adLockReadOnly

CerrarRecorset rstUniversal

LstGuias.ListItems.Clear
For II = FufuLo To FufuLo2 Step 1
  Set Item = LstGuias.ListItems.Add(, , II)
Next
End Sub
