VERSION 5.00
Begin VB.Form FrmVerRelEntregaDoc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ver relacion de entrega documento..."
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdVerGuias 
      Caption         =   "Ver guias"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8535
      Begin VB.TextBox TxtCampo 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TxtCampo 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtNmTercero 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2880
         MaxLength       =   60
         TabIndex        =   4
         ToolTipText     =   "Aqui se debe ingresar el tercero"
         Top             =   600
         Width           =   5535
      End
      Begin VB.TextBox TxtCampo 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   1080
         MaxLength       =   11
         TabIndex        =   3
         ToolTipText     =   "Aqui se debe ingresar el tercero"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox TxtCampo 
         Height          =   1125
         Index           =   3
         Left            =   1080
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   7335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   2280
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         Height          =   195
         Left            =   690
         TabIndex        =   9
         Top             =   240
         Width           =   210
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Tercero:"
         Height          =   195
         Index           =   33
         Left            =   300
         TabIndex        =   8
         Top             =   600
         Width           =   600
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Comentario:"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   960
         Width           =   840
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   255
      Left            =   6720
      TabIndex        =   0
      Top             =   2520
      Width           =   1935
   End
End
Attribute VB_Name = "FrmVerRelEntregaDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
  Unload Me
End Sub

Private Sub CmdVerGuias_Click()
  FufuLo = Val(TxtCampo(0).Text)
  FufuSt = "S"
  FrmLlenarRelEntegaDoc.Show 1
End Sub

Private Sub Form_Load()
  AbrirRecorset rstUniversal, "SELECT relentregadoc.*, terceros.RazonSocial FROM relentregadoc LEFT JOIN terceros ON relentregadoc.IdTercero = terceros.IDTercero WHERE (((relentregadoc.IDRel)=" & FufuLo & "));", CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.EOF = False Then
    TxtCampo(0).Text = rstUniversal.Fields("IdRel") & ""
    TxtCampo(1).Text = Format(rstUniversal.Fields("Fecha") & "", "dd/mm/yy")
    TxtCampo(2).Text = rstUniversal.Fields("IdTercero") & ""
    TxtCampo(3).Text = rstUniversal.Fields("Comentarios") & ""
    TxtNmTercero.Text = rstUniversal.Fields("RazonSocial") & ""
  End If
  CerrarRecorset rstUniversal
End Sub
