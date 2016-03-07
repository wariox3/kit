VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmInformeRecibosCaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametros recibos caja"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   6720
      TabIndex        =   15
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton OptResumen 
         Caption         =   "Resumen"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OpDetalle 
         Caption         =   "Detalle"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.CheckBox ChkFiltrarFecha 
      Caption         =   "Filtrar por fecha de pago"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   1440
      Width           =   2415
   End
   Begin VB.ComboBox CboTpRecibo 
      Height          =   315
      ItemData        =   "FrmInformeRecibosCaja.frx":0000
      Left            =   1680
      List            =   "FrmInformeRecibosCaja.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton CmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox TxtIdTercero 
      Height          =   315
      Left            =   1665
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox TxtNumero 
      Height          =   315
      Left            =   1665
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.CheckBox ChkExcel 
      Caption         =   "Excel"
      Height          =   255
      Left            =   7305
      TabIndex        =   0
      Top             =   1320
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DPFechaDesde 
      Height          =   300
      Left            =   1680
      TabIndex        =   10
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16842753
      CurrentDate     =   38971
   End
   Begin MSComCtl2.DTPicker DPFechaHasta 
      Height          =   300
      Left            =   1680
      TabIndex        =   11
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16842753
      CurrentDate     =   38971
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   1080
      TabIndex        =   13
      Top             =   1800
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Desde:"
      Height          =   195
      Left            =   1080
      TabIndex        =   12
      Top             =   1440
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      Height          =   195
      Left            =   1185
      TabIndex        =   8
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tercero:"
      Height          =   195
      Left            =   945
      TabIndex        =   7
      Top             =   960
      Width           =   600
   End
   Begin VB.Label LblNmTercero 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3225
      TabIndex        =   6
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Numero:"
      Height          =   195
      Left            =   945
      TabIndex        =   5
      Top             =   600
      Width           =   600
   End
End
Attribute VB_Name = "FrmInformeRecibosCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CboTpRecibo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub CmdGenerar_Click()
  Dim strSql As String
  Dim strEntidad As String
  If OptResumen.Value = True Then
    strEntidad = "sql_ic_recibos"
    varParametrosRecibo.InformeDetallado = False
  Else
    strEntidad = "sql_ic_recibos_detalles"
    varParametrosRecibo.InformeDetallado = True
  End If
  varParametrosRecibo.Generar = True
  If Val(TxtIdTercero.Text) > 0 Then
    varParametrosRecibo.IdCliente = Val(TxtIdTercero.Text)
  End If
  If Val(TxtNumero.Text) > 0 Then
    varParametrosRecibo.Numero = Val(TxtNumero.Text)
  End If
  If CboTpRecibo.ListIndex > 0 Then
    varParametrosRecibo.Tipo = CboTpRecibo.ListIndex
  End If
  If ChkFiltrarFecha.Value = 1 Then
    varParametrosRecibo.Fecha = True
    varParametrosRecibo.FechaDesde = Format(DPFechaDesde.Value, "yyyy/mm/dd")
    varParametrosRecibo.FechaHasta = Format(DPFechaHasta.Value, "yyyy/mm/dd")
  End If
  
  strSql = "Select " & strEntidad & ".* FROM " & strEntidad & " WHERE 1"
  
  If varParametrosRecibo.IdCliente <> 0 Then
    strSql = strSql & " AND IdTercero = " & varParametrosRecibo.IdCliente
  End If
  If varParametrosRecibo.Tipo <> 0 Then
    strSql = strSql & " AND IdReciboTipo = " & varParametrosRecibo.Tipo
  End If
  If varParametrosRecibo.Numero <> 0 Then
    strSql = strSql & " AND numero = " & varParametrosRecibo.Numero
  End If
  If varParametrosRecibo.Fecha = True Then
    strSql = strSql & " AND (FechaPago >= '" & varParametrosRecibo.FechaDesde & "' AND FechaPago <= '" & varParametrosRecibo.FechaHasta & "')"
  End If
  If ChkExcel.Value = 1 Then
    Dim rstRecibosCaja As New ADODB.Recordset
    rstRecibosCaja.CursorLocation = adUseClient
    AbrirRecorset rstRecibosCaja, strSql, CnnPrincipal, adOpenDynamic, adLockReadOnly
    If rstRecibosCaja.State = adStateOpen Then
      If rstRecibosCaja.EOF = False Then
        ExportarExcel rstRecibosCaja
        MsgBox "Se ha exportado con exito", vbInformation
      End If
    End If
    varParametrosRecibo.Generar = False
  End If
  varParametrosRecibo.sql = strSql
  Unload Me
End Sub

Private Sub Form_Load()
  varParametrosRecibo.Generar = False
  varParametrosRecibo.IdCliente = 0
  varParametrosRecibo.Tipo = 0
  varParametrosRecibo.GenerarExcel = False
  varParametrosRecibo.Numero = 0
  varParametrosRecibo.Fecha = False
  varParametrosRecibo.FechaDesde = ""
  varParametrosRecibo.FechaHasta = ""
  CboTpRecibo.ListIndex = 0
  DPFechaDesde.Value = Date
  DPFechaHasta.Value = Date
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

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

