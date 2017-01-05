VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmInformeNotaCredito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe nota credito"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkExcel 
      Caption         =   "Excel"
      Height          =   255
      Left            =   6600
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox TxtNumero 
      Height          =   315
      Left            =   960
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox TxtIdTercero 
      Height          =   315
      Left            =   960
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton CmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   6135
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4695
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox CboTpNotaCredito 
      Height          =   315
      ItemData        =   "FrmInformeNotaCredito.frx":0000
      Left            =   975
      List            =   "FrmInformeNotaCredito.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   4800
   End
   Begin VB.CheckBox ChkFiltrarFecha 
      Caption         =   "Filtrar por fecha de pago"
      Height          =   255
      Left            =   2535
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   6015
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton OpDetalle 
         Caption         =   "Detalle"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton OptResumen 
         Caption         =   "Resumen"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin MSComCtl2.DTPicker DPFechaDesde 
      Height          =   300
      Left            =   975
      TabIndex        =   10
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16777217
      CurrentDate     =   38971
   End
   Begin MSComCtl2.DTPicker DPFechaHasta 
      Height          =   300
      Left            =   975
      TabIndex        =   11
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16777217
      CurrentDate     =   38971
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Numero:"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   600
      Width           =   600
   End
   Begin VB.Label LblNmTercero 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2520
      TabIndex        =   16
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tercero:"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   960
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      Height          =   195
      Left            =   480
      TabIndex        =   14
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Desde:"
      Height          =   195
      Left            =   375
      TabIndex        =   13
      Top             =   1440
      Width           =   510
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   375
      TabIndex        =   12
      Top             =   1800
      Width           =   465
   End
End
Attribute VB_Name = "FrmInformeNotaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub CboTpNotaCredito_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub CmdGenerar_Click()
  Dim strSql As String
  Dim strEntidad As String
  If OptResumen.Value = True Then
    strEntidad = "sql_ic_nota_credito"
    varParametrosNotaCredito.InformeDetallado = False
  Else
    strEntidad = "sql_ic_nota_credito_detalles"
    varParametrosNotaCredito.InformeDetallado = True
  End If
  varParametrosNotaCredito.Generar = True
  If Val(TxtIdTercero.Text) > 0 Then
    varParametrosNotaCredito.IdCliente = Val(TxtIdTercero.Text)
  End If
  If Val(TxtNumero.Text) > 0 Then
    varParametrosNotaCredito.Numero = Val(TxtNumero.Text)
  End If
  If CboTpNotaCredito.ListIndex > 0 Then
    varParametrosNotaCredito.Tipo = CboTpNotaCredito.ListIndex
  End If
  If ChkFiltrarFecha.Value = 1 Then
    varParametrosNotaCredito.Fecha = True
    varParametrosNotaCredito.FechaDesde = Format(DPFechaDesde.Value, "yyyy/mm/dd")
    varParametrosNotaCredito.FechaHasta = Format(DPFechaHasta.Value, "yyyy/mm/dd")
  End If
  
  strSql = "Select " & strEntidad & ".* FROM " & strEntidad & " WHERE 1"
  
  If varParametrosNotaCredito.IdCliente <> 0 Then
    strSql = strSql & " AND IdTercero = " & varParametrosNotaCredito.IdCliente
  End If
  If varParametrosNotaCredito.Tipo <> 0 Then
    strSql = strSql & " AND IdNotaCreditoTipo = " & varParametrosNotaCredito.Tipo
  End If
  If varParametrosNotaCredito.Numero <> 0 Then
    strSql = strSql & " AND numero = " & varParametrosNotaCredito.Numero
  End If
  If varParametrosNotaCredito.Fecha = True Then
    strSql = strSql & " AND (Fecha >= '" & varParametrosNotaCredito.FechaDesde & "' AND Fecha <= '" & varParametrosNotaCredito.FechaHasta & "')"
  End If
  If ChkExcel.Value = 1 Then
    Dim rstNotaCredito As New ADODB.Recordset
    rstNotaCredito.CursorLocation = adUseClient
    AbrirRecorset rstNotaCredito, strSql, CnnPrincipal, adOpenDynamic, adLockReadOnly
    If rstNotaCredito.State = adStateOpen Then
      If rstNotaCredito.EOF = False Then
        ExportarExcel rstNotaCredito
        MsgBox "Se ha exportado con exito", vbInformation
      End If
    End If
    varParametrosNotaCredito.Generar = False
  End If
  varParametrosNotaCredito.sql = strSql
  Unload Me
End Sub

Private Sub Form_Load()
  varParametrosNotaCredito.Generar = False
  varParametrosNotaCredito.IdCliente = 0
  varParametrosNotaCredito.Tipo = 0
  varParametrosNotaCredito.GenerarExcel = False
  varParametrosNotaCredito.Numero = 0
  varParametrosNotaCredito.Fecha = False
  varParametrosNotaCredito.FechaDesde = ""
  varParametrosNotaCredito.FechaHasta = ""
  CboTpNotaCredito.ListIndex = 0
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



