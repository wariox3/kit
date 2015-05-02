VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmInformeNovedades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver informe novedades"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkFechas 
      Caption         =   "Filtrar por fecha"
      Height          =   255
      Left            =   5280
      TabIndex        =   16
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton CmdVerInforme 
      Caption         =   "Ver informe"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox TxtIdTercero 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Width           =   1695
      Begin VB.OptionButton OptTodas 
         Caption         =   "Todas"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton OptPendientes 
         Caption         =   "Pendientes"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptSolucionadas 
         Caption         =   "Solucionadas"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox TxtNovedad 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   5535
   End
   Begin VB.TextBox TxtIdNovedad 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTPFechaDesde 
      Height          =   315
      Left            =   5280
      TabIndex        =   12
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yy"
      Format          =   49807363
      CurrentDate     =   39740
   End
   Begin MSComCtl2.DTPicker DTPFechaHasta 
      Height          =   315
      Left            =   5280
      TabIndex        =   13
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yy"
      Format          =   49807363
      CurrentDate     =   39740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   4680
      TabIndex        =   15
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label LblFecha 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   195
      Left            =   4680
      TabIndex        =   14
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tercero:"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   600
   End
   Begin VB.Label LblNmTercero 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label LblIdNov 
      AutoSize        =   -1  'True
      Caption         =   "Novedad:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   705
   End
End
Attribute VB_Name = "FrmInformeNovedades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub CmdVerInforme_Click()
  Dim strSql As String
  strSql = "SELECT sql_im_novedades_cliente.* FROM sql_im_novedades_cliente WHERE 1 "
  If OptSolucionadas.value = True Then
    strSql = strSql & " AND Solucionada = 1"
  End If
  If OptPendientes.value = True Then
    strSql = strSql & " AND Solucionada = 0"
  End If
  If Val(TxtIdNovedad.Text) <> 0 Then
    strSql = strSql & " AND IdNovedad = " & Val(TxtIdNovedad.Text)
  End If
  If Val(TxtIdTercero.Text) <> 0 Then
    strSql = strSql & " AND Cuenta = '" & TxtIdTercero.Text & "'"
  End If
  If ChkFechas.value = 1 Then
    strSql = strSql & " AND FhEntradaBodega >='" & Format(DTPFechaDesde.value, "yyyy/mm/dd") & " 00:00:00' AND FhEntradaBodega <='" & Format(DTPFechaHasta.value, "yyyy/mm/dd") & " 23:59:00' "
  End If
  Mostrar_Reporte CnnPrincipal, 36, strSql, "Novedades", 2
End Sub


Private Sub Form_Load()
  DTPFechaDesde.value = Date
  DTPFechaHasta.value = Date
End Sub

Private Sub TxtIdNovedad_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirConsultaGral "IdNovedad", "NmNovedad", "CausalesNovedad", CnnPrincipal
    TxtIdNovedad.Text = Principal.ToolConsultas1.DatLo
  End If
End Sub

Private Sub TxtIdNovedad_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdNovedad_Validate(Cancel As Boolean)
  AbrirRecorset rstUniversal, "select IdNovedad, NmNovedad from CausalesNovedad where IdNovedad=" & Val(TxtIdNovedad), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = False Then
    TxtNovedad = rstUniversal!NmNovedad
  Else
    TxtIdNovedad.Text = "": TxtNovedad.Text = ""
  End If
  CerrarRecorset rstUniversal
End Sub

Private Sub TxtIdTercero_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsulta 7, CnnPrincipal
    TxtIdTercero.Text = Principal.ToolConsultas1.DatSt
  End If
End Sub

Private Sub TxtIdTercero_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
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
