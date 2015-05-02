VERSION 5.00
Begin VB.Form FrmInformeCarteraEdades 
   Caption         =   "Parametros informe"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtIdTercero 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.ComboBox CboTipo 
      Height          =   315
      ItemData        =   "FrmInformeCarteraEdades.frx":0000
      Left            =   1080
      List            =   "FrmInformeCarteraEdades.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton CmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label LblNmTercero 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tercero:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "FrmInformeCarteraEdades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub CmdGenerar_Click()
  Dim strWhere As String
  strWhere = " WHERE 1 "
  If CboTipo.ListIndex > 0 Then
    strWhere = strWhere & " AND TipoFactura = " & CboTipo.ListIndex
  End If
  If Val(TxtIdTercero.Text) <> 0 Then
    strWhere = strWhere & " AND IdTercero = '" & TxtIdTercero.Text & "'"
  End If
  Mostrar_Reporte CnnPrincipal, 30, "SELECT sql_ic_cartera_edades.* FROM sql_ic_cartera_edades " & strWhere, "Cartera por edades", 2
End Sub

Private Sub TxtIdTercero_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsulta 7, CnnPrincipal
    TxtIdTercero.Text = Principal.ToolConsultas1.DatSt
  End If
End Sub

Private Sub TxtIdTercero_LostFocus()
  If TxtIdTercero = "0" And TxtIdTercero.Text = "" Then
      LblNmTercero.Caption = ""
  Else
    AbrirRecorset rstUniversal, "SELECT IdTercero, RazonSocial " & _
                                "FROM terceros " & _
                                "WHERE IdTercero='" & TxtIdTercero.Text & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      LblNmTercero.Caption = rstUniversal!RazonSocial & ""
    Else
      LblNmTercero.Caption = "": TxtIdTercero.Text = ""
    End If
    CerrarRecorset rstUniversal
  End If
End Sub
