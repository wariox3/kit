VERSION 5.00
Begin VB.Form FrmSolucionarNovedad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Solucionar novedad"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Frame FraSolucion 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   10215
      Begin VB.TextBox TxtSolucion 
         Height          =   1455
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   9255
      End
      Begin VB.Label LblNotas 
         AutoSize        =   -1  'True
         Caption         =   "Solucion:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Label LblIdNovedad 
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ID Novedad:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmSolucionarNovedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub CmdGuardar_Click()
      If MsgBox("Va a solucionar la novedad con:" & Chr(13) & TxtSolucion.Text & Chr(13) & "¿Esta seguro que dese solucionar la novedad [" & LblIdNovedad.Caption & "] ahora?", vbQuestion + vbYesNo) = vbYes Then
        AbrirRecorset rstUniversal, "UPDATE novedades SET Solucionada = 1, Solucion='" & TxtSolucion & "', FHSolucion='" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', UsuSol='" & CodUsuarioActivo & "' where ID=" & LblIdNovedad.Caption, CnnPrincipal, adOpenDynamic, adLockOptimistic
        Unload Me
      End If
End Sub

Private Sub Form_Load()
  LblIdNovedad.Caption = FufuLo
End Sub
