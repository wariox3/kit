VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAgregarCuentaCobrarMasivo 
   Caption         =   "Agregar cuenta cobrar masivo"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   11805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkRetencionFuenteSinBase 
      Caption         =   "Retencion fuente sin base"
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   6000
      Width           =   2775
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   9960
      TabIndex        =   3
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   1815
   End
   Begin MSComctlLib.ListView LstCuentaCobrar 
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipo"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Numero"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Saldo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Ica"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label LlbNmTercero 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label LblIdRecibo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label LblNit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FrmAgregarCuentaCobrarMasivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCuentasCobrar As New ADODB.Recordset
Dim RetencionFuenteCliente As Integer
Private Sub Ver()
  Dim Consulta As String
  Dim rstCuentaCobrar As New ADODB.Recordset
  rstCuentaCobrar.CursorLocation = adUseClient
  Consulta = "SELECT cuentas_cobrar.*, NmTipoFactura FROM cuentas_cobrar LEFT JOIN facturas_tipos ON cuentas_cobrar.TipoFactura=facturas_tipos.IdTipoFactura WHERE IdTercero = '" & Val(LblNit.Caption) & "' AND saldo > 0 order by NroDocumento"
  Consulta = "SELECT cuentas_cobrar.* FROM cuentas_cobrar WHERE IdTercero = '" & Val(LblNit.Caption) & "' AND saldo > 0 order by NroDocumento"
  AbrirRecorset rstCuentaCobrar, Consulta, CnnPrincipal, adOpenStatic, adLockReadOnly
  
  LstCuentaCobrar.ListItems.Clear
  If rstCuentaCobrar.RecordCount > 0 Then
    Do While rstCuentaCobrar.EOF = False
      Set Item = LstCuentaCobrar.ListItems.Add(, , rstCuentaCobrar!IdCxC)
      'Item.SubItems(1) = rstCuentaCobrar!NmTipoFactura & ""
      Item.SubItems(2) = rstCuentaCobrar!NroDocumento & ""
      Item.SubItems(3) = rstCuentaCobrar!FechaDoc & ""
      Item.SubItems(4) = rstCuentaCobrar!saldo & ""
      rstCuentaCobrar.MoveNext
    Loop
  End If
  
  'Set GrillaCuentasCobrar.DataSource = rstCuentaCobrar
End Sub

Private Sub CmdAgregar_Click()
  Dim rstCuentaCobrar As New ADODB.Recordset
  rstCuentaCobrar.CursorLocation = adUseClient
  Dim total As Double
  Dim saldo As Double
  Dim retencionFuente As Double
  II = 1
  While II <= LstCuentaCobrar.ListItems.Count
    If LstCuentaCobrar.ListItems(II).Checked = True Then
      AbrirRecorset rstCuentaCobrar, "SELECT cuentas_cobrar.* FROM cuentas_cobrar WHERE IdCxC = " & Val(LstCuentaCobrar.ListItems(II)), CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstCuentaCobrar.RecordCount > 0 Then
        saldo = rstCuentaCobrar.Fields("Saldo")
        total = rstCuentaCobrar.Fields("Saldo")
        retencionFuente = 0
        If RetencionFuenteCliente = 1 Then
          If rstCuentaCobrar.Fields("Saldo") > 127000 Or ChkRetencionFuenteSinBase.Value = 1 Then
            retencionFuente = Round(saldo * 1 / 100)
          End If
        End If
        total = Round(total - retencionFuente)
        If total <= rstCuentaCobrar!saldo Then
          AbrirRecorset rstUniversal, "INSERT INTO recibos_caja_det (IdRecibo, codigo_cuenta_cobrar_fk, valor, descuento, ajuste_peso, retencion_ica, retencion_fuente) VALUES (" & Val(LblIdRecibo.Caption) & ", " & rstCuentaCobrar.Fields("IdCxC") & ", " & total & ", 0, 0, 0, " & retencionFuente & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
          AbrirRecorset rstUniversal, "UPDATE recibos_caja SET Total = Total + " & total & ", descuento = descuento + 0, ajuste_peso = ajuste_peso + 0, retencion_ica = retencion_ica + 0, retencion_fuente = retencion_fuente + " & retencionFuente & " WHERE IdRecibo = " & Val(LblIdRecibo.Caption), CnnPrincipal, adOpenDynamic, adLockOptimistic
          AbrirRecorset rstUniversal, "UPDATE cuentas_cobrar SET Saldo = Saldo - " & saldo & " WHERE IdCxC = " & rstCuentaCobrar!IdCxC, CnnPrincipal, adOpenDynamic, adLockOptimistic
        End If
      End If
      CerrarRecorset rstCuentaCobrar
      LstCuentaCobrar.ListItems.Remove (II)
    Else
      II = II + 1
    End If
  Wend
  Unload Me
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim rstTercero As New ADODB.Recordset
  rstTercero.CursorLocation = adUseClient
  rstCuentasCobrar.CursorLocation = adUseClient
  LblIdRecibo.Caption = FufuLo
  LblNit.Caption = FufuSt
  AbrirRecorset rstTercero, "SELECT RetencionFuente, RazonSocial FROM terceros WHERE IDTercero = '" & FufuSt & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
    RetencionFuenteCliente = Val(rstTercero.Fields("RetencionFuente"))
    LlbNmTercero.Caption = rstTercero.Fields("RazonSocial") & ""
  CerrarRecorset rstTercero
  Ver
End Sub

