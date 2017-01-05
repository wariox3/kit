VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRegistroDeMonitoreo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de monitoreo..."
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   5280
      TabIndex        =   10
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox TxtNotas 
      Height          =   975
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1320
      Width           =   7335
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Format          =   16777217
      CurrentDate     =   39146
   End
   Begin MSComCtl2.DTPicker DTPHora 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Format          =   16777218
      UpDown          =   -1  'True
      CurrentDate     =   39146
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   7080
      TabIndex        =   9
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox TxtNmControlPost 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   240
      Width           =   6615
   End
   Begin VB.TextBox TxtIdControlPost 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
   Begin VB.Label LblDespacho 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   1080
      TabIndex        =   12
      Top             =   2400
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Monitoreo:"
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
      TabIndex        =   11
      Top             =   2400
      Width           =   915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Id Control Post:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   195
      Left            =   840
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Hora:"
      Height          =   195
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Notas:"
      Height          =   195
      Left            =   840
      TabIndex        =   4
      Top             =   1320
      Width           =   465
   End
End
Attribute VB_Name = "FrmRegistroDeMonitoreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
  If TxtIdControlPost.Text <> "" Then
    rstUniversal.Open "INSERT INTO monitoreocontrolpost (IdMonitoreo, IdControlPost, FhHrReporte, Notas, usuario) VALUES (" & Val(LblDespacho) & ", " & Val(TxtIdControlPost.Text) & ", '" & Format(DTPFecha.Value, "yyyy/mm/dd") & " " & Format(DTPHora.Value, "h:m:s") & "', '" & TxtNotas & "', '" & NmUsuarioActivo & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
    rstUniversal.Open "Update MonitoreoVehiculos set UltReporte='" & Format(DTPFecha.Value, "yyyy/mm/dd") & " " & Format(DTPHora.Value, "h:m:s") & "' where ID=" & LblDespacho, CnnPrincipal, adOpenDynamic, adLockOptimistic
    MsgBox "Monitoreo ingresado con exito", vbInformation
    Unload Me
  Else
    MsgBox "Debe seleccionar un Control Post para el reporte", vbCritical: TxtIdControlPost.SetFocus
  End If
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub
Private Sub DTPFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then DTPHora.SetFocus
End Sub
Private Sub DTPHora_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then TxtNotas.SetFocus
End Sub

Private Sub Form_Load()
  DTPFecha.Value = Date
  DTPHora.Value = Time
  LblDespacho.Caption = FufuLo
End Sub

Private Sub TxtIdControlPost_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
      Principal.ToolConsultas1.AbrirDevConsulta 6, CnnPrincipal
      TxtIdControlPost.Text = Principal.ToolConsultas1.DatLo
  End If
End Sub
Private Sub TxtIdControlPost_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub TxtIdControlPost_LostFocus()
  If Val(TxtIdControlPost.Text) <> 0 Then
    rstUniversal.Open "SELECT IdControlPost, NmControlPost From ControlPost where IdControlPost=" & TxtIdControlPost, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      TxtNmControlPost = rstUniversal!NmControlPost & ""
    Else
      TxtNmControlPost.Text = "": TxtIdControlPost.Text = ""
    End If
    rstUniversal.Close
  End If
End Sub
Private Sub TxtNotas_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys vbTab
  End If
End Sub
