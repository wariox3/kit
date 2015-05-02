VERSION 5.00
Begin VB.Form FrmInformacionLiquidarDespacho 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Liquidar despacho"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1515
         TabIndex        =   12
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton CmdCalcular 
         Caption         =   "Calcular"
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox TxtPorcentaje 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox TxtManejoCE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TxtFleteCE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox TxtManejo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox TxtFlete 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Flete conductor:"
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   2520
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "Porcentaje:"
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Manejo CE:"
         Height          =   195
         Left            =   525
         TabIndex        =   8
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Flete CE:"
         Height          =   195
         Left            =   705
         TabIndex        =   7
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Manejo:"
         Height          =   195
         Left            =   780
         TabIndex        =   6
         Top             =   720
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Flete:"
         Height          =   195
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   390
      End
   End
End
Attribute VB_Name = "FrmInformacionLiquidarDespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim rstDespacho As New ADODB.Recordset
AbrirRecorset rstDespacho, "SELECT despachos.* FROM despachos WHERE OrdDespacho = " & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
  TxtFlete.Text = Format(rstDespacho!FleteCobra, "#,##0.00;(#,##0.00)")
  TxtManejo.Text = Format(rstDespacho!ManejoCobra, "#,##0.00;(#,##0.00)")
  TxtFleteCE.Text = Format(rstDespacho!FleteCE, "#,##0.00;(#,##0.00)")
  TxtManejoCE.Text = Format(rstDespacho!ManejoCE, "#,##0.00;(#,##0.00)")
CerrarRecorset rstDespacho
End Sub

