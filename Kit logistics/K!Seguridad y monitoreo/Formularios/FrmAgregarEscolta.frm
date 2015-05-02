VERSION 5.00
Begin VB.Form FrmAgregarEscolta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar acompañamiento..."
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtComentarios 
      Height          =   855
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   6135
   End
   Begin VB.TextBox TxtVrEscoltada 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox TxtNmEscolta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   6135
   End
   Begin VB.TextBox TxtIdEscolta 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton CmdVrAgregarAcompañamiento 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label LblIdMonitoreo 
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label4 
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
      TabIndex        =   8
      Top             =   1920
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Comentarios:"
      Height          =   195
      Left            =   840
      TabIndex        =   7
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Vr acompañamiento:"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   210
   End
End
Attribute VB_Name = "FrmAgregarEscolta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstAgregarAcompañamiento As New ADODB.Recordset


Private Sub CmdVrAgregarAcompañamiento_Click()
  If TxtIdEscolta.Text <> "" Then
    rstAgregarAcompañamiento.Open "insert into monitoreo_acompañamiento (IdMonitoreo, IdEscolta, VrAcompañamiento, ComentariosAcompañamiento) values(" & Val(LblIdMonitoreo.Caption) & ", " & Val(TxtIdEscolta.Text) & ", " & Val(TxtVrEscoltada.Text) & ", '" & TxtComentarios.Text & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
    Unload Me
  Else
    MsgBox "No a especificado la persona que realiza el acompañamiento", vbCritical, "Error al agregar acompañamiento"
  End If
End Sub

Private Sub Form_Load()
  rstAgregarAcompañamiento.CursorLocation = adUseClient
  LblIdMonitoreo.Caption = FufuLo
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstAgregarAcompañamiento = Nothing
End Sub

Private Sub TxtComentarios_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdEscolta_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsulta 7, CnnPrincipal
    TxtIdEscolta.Text = Principal.ToolConsultas1.DatSt
  End If
End Sub

Private Sub TxtIdEscolta_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtIdEscolta, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdEscolta_Validate(Cancel As Boolean)
  AbrirRecorset rstUniversal, "Select IdTercero, RazonSocial, IdCliente from Terceros where IdTercero='" & TxtIdEscolta.Text & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = False Then
    TxtNmEscolta.Text = rstUniversal.Fields("RazonSocial") & ""
  Else
    TxtNmEscolta.Text = "": TxtIdEscolta.Text = ""
  End If
  CerrarRecorset rstUniversal
End Sub

Private Sub TxtVrEscoltada_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtVrEscoltada, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
