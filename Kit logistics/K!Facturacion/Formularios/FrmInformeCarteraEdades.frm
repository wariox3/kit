VERSION 5.00
Begin VB.Form FrmInformeCarteraEdades 
   Caption         =   "Parametros informe"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkExcel 
      Caption         =   "Excel"
      Height          =   255
      Left            =   7320
      TabIndex        =   15
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox TxtNumero 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox TxtIdCentroOperaciones 
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox TxtIdAsesor 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox TxtNmAsesor 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3240
      TabIndex        =   10
      Top             =   1320
      Width           =   4935
   End
   Begin VB.TextBox TxtIdTercero 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.ComboBox CboTipo 
      Height          =   315
      ItemData        =   "FrmInformeCarteraEdades.frx":0000
      Left            =   1680
      List            =   "FrmInformeCarteraEdades.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton CmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Numero:"
      Height          =   195
      Left            =   960
      TabIndex        =   14
      Top             =   600
      Width           =   600
   End
   Begin VB.Label LblNmCentroOperaciones 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3240
      TabIndex        =   13
      Top             =   1680
      Width           =   4935
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Centro operaciones:"
      Height          =   195
      Left            =   135
      TabIndex        =   12
      Top             =   1680
      Width           =   1425
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Asesor:"
      Height          =   195
      Left            =   1035
      TabIndex        =   11
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label LblNmTercero 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3240
      TabIndex        =   9
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tercero:"
      Height          =   195
      Left            =   960
      TabIndex        =   8
      Top             =   960
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      Height          =   195
      Left            =   1200
      TabIndex        =   7
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

Private Sub CboTipo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub CmdGenerar_Click()
  Dim strSql As String
  varParametrosCartera.Generar = True
  If Val(TxtIdTercero.Text) > 0 Then
    varParametrosCartera.IdCliente = Val(TxtIdTercero.Text)
  End If
  If Val(TxtIdAsesor.Text) > 0 Then
    varParametrosCartera.IdAsesor = Val(TxtIdAsesor.Text)
  End If
  If Val(TxtNumero.Text) > 0 Then
    varParametrosCartera.numero = Val(TxtNumero.Text)
  End If
  If CboTipo.ListIndex > 0 Then
    varParametrosCartera.Tipo = CboTipo.ListIndex
  End If
  
  strSql = "Select sql_ic_cartera_edades.* FROM sql_ic_cartera_edades WHERE 1"
  
  If varParametrosCartera.IdAsesor <> 0 Then
    strSql = strSql & " AND IdAsesor = " & varParametrosCartera.IdAsesor
  End If
  If varParametrosCartera.IdCliente <> 0 Then
    strSql = strSql & " AND IdTercero = " & varParametrosCartera.IdCliente
  End If
  If varParametrosCartera.IdCentroOperaciones <> 0 Then
    strSql = strSql & " AND IdPO = " & varParametrosCartera.IdCentroOperaciones
  End If
  If varParametrosCartera.Tipo <> 0 Then
    strSql = strSql & " AND TipoFactura = " & varParametrosCartera.Tipo
  End If
  If varParametrosCartera.numero <> 0 Then
    strSql = strSql & " AND NroDocumento = " & varParametrosCartera.numero
  End If
  If ChkExcel.Value = 1 Then
    Dim rstCuentasCobrar As New ADODB.Recordset
    rstCuentasCobrar.CursorLocation = adUseClient
    AbrirRecorset rstCuentasCobrar, strSql, CnnPrincipal, adOpenDynamic, adLockReadOnly
    If rstCuentasCobrar.State = adStateOpen Then
      If rstCuentasCobrar.EOF = False Then
        ExportarExcel rstCuentasCobrar
        MsgBox "Se ha exportado con exito", vbInformation
      End If
    End If
    varParametrosCartera.Generar = False
  End If
  varParametrosCartera.sql = strSql
  Unload Me
End Sub

Private Sub Form_Load()
  varParametrosCartera.Generar = False
  varParametrosCartera.IdCliente = 0
  varParametrosCartera.IdAsesor = 0
  varParametrosCartera.IdCentroOperaciones = 0
  varParametrosCartera.Tipo = 0
  varParametrosCartera.GenerarExcel = False
  varParametrosCartera.numero = 0
  CboTipo.ListIndex = 0
End Sub

Private Sub TxtIdAsesor_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    FrmBuscarAsesor.Show 1
    TxtIdAsesor.Text = FufuLo
  End If
End Sub

Private Sub TxtIdAsesor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdAsesor_Validate(Cancel As Boolean)
    If Val(TxtIdAsesor.Text) <> 0 Then
      AbrirRecorset rstUniversal, "SELECT IdAsesor, NmAsesor From Asesores where IdAsesor=" & Val(TxtIdAsesor.Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        TxtNmAsesor.Text = rstUniversal!NmAsesor & ""
      Else
        TxtNmAsesor.Text = "": TxtIdAsesor.Text = ""
      End If
      CerrarRecorset rstUniversal
    Else
      TxtIdAsesor.Text = ""
      TxtNmAsesor.Text = ""
    End If
End Sub

Private Sub TxtIdCentroOperaciones_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    FrmBuscarCO.Show 1
    TxtIdCentroOperaciones.Text = FufuLo
  End If
End Sub

Private Sub TxtIdCentroOperaciones_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdCentroOperaciones_Validate(Cancel As Boolean)
    If Val(TxtIdCentroOperaciones.Text) <> 0 Then
      AbrirRecorset rstUniversal, "SELECT IDPO, NmPuntoOperaciones From centrosoperaciones where IDPO=" & TxtIdCentroOperaciones.Text, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        LblNmCentroOperaciones.Caption = rstUniversal!NmPuntoOperaciones & ""
      Else
        LblNmCentroOperaciones.Caption = "": TxtIdCentroOperaciones.Text = ""
      End If
      CerrarRecorset rstUniversal
    Else
        LblNmCentroOperaciones.Caption = "": TxtIdCentroOperaciones.Text = ""
    End If
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
