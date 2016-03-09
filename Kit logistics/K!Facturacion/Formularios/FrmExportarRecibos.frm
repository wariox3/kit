VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmExportarRecibos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar recibos caja"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   11340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExportarSiigoCotrascal 
      Caption         =   "Exportar SIIGO"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   6120
      Width           =   1815
   End
   Begin VB.TextBox TxtRuta 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   6015
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   9840
      TabIndex        =   0
      Top             =   6120
      Width           =   1335
   End
   Begin MSComctlLib.ListView LstRecibos 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9763
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Numero"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tercero"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label LblRuta 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   390
   End
End
Attribute VB_Name = "FrmExportarRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstRecibosExp As New ADODB.Recordset

Private Sub CmdExportarSiigoCotrascal_Click()
  Dim RutaSalida As String
  Dim Fila        As Long
  Dim Columna     As Long
  
On Error GoTo Error_Handler
    RutaSalida = TxtRuta.Text & "recexpsiigo" & Format(Date, "dd.mm.yy") & "." & Format(Time, "hh.mm.ss") & ".txt"
    Dim J As Integer
    Dim strCuenta As String
    Dim strDetalle As String
    Dim strTipo As String
    Dim strComprobante As String
    Dim strNit As String
    Dim strCentroCostos As String
    Dim strVendedor As String
    Dim douValor As Double
    Dim douRetencionFuente As Double
    Dim strValor As String
    Dim strNumero As String
    Dim intNroRegistros As Integer
    Dim strDocumentoCruce As String
    Fila = 0
    II = 1
    Open RutaSalida For Append As #1
    'Print #1, "Cuenta  Comprobante Fecha(mm/dd/yyyy) Documento Documento Ref.  Nit Detalle Tipo  Valor Base  Centro de Costo Trans. Ext  Plazo"
    While II <= LstRecibos.ListItems.Count
      If LstRecibos.ListItems(II).Checked = True Then
        rstRecibosExp.Open "SELECT recibos_caja.*, terceros.RazonSocial " & _
                            "FROM recibos_caja " & _
                            "LEFT JOIN terceros ON recibos_caja.IdTercero = terceros.IdTercero " & _
                            "WHERE Exportado=0 AND IdRecibo = " & LstRecibos.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic

        Print #1, "R" & strComprobante & strNumero & Rellenar(J & "", 5, "0", 1) & Rellenar(strNit, 13, "0", 1) & "000" & strCuenta & "000000000000000" & Format(rstRecibosExp!Fecha, "yyyymmdd") & strCentroCostos & "000" & Rellenar(strDetalle, 50, " ", 0) & strTipo & Rellenar(strValor, 15, "0", 1) & "000000000000000" & strVendedor & "0001" & "001" & "0001" & "000" & "000000000000000" & strDocumentoCruce
     
        rstRecibosExp.Close
        'rstRecibosExp.Open "UPDATE facturas_venta SET Exportada=1 where Numero=" & LstRecibos.ListItems(II) & " AND TipoFactura = " & LstRecibos.ListItems(II).SubItems(1), CnnPrincipal, adOpenDynamic, adLockOptimistic
        LstRecibos.ListItems.Remove (II)
      Else
       II = II + 1
      End If
    Wend
    Close #1
  
  Exit Sub
Error_Handler:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  rstRecibosExp.CursorLocation = adUseClient
  TxtRuta.Text = GetSetting("Kit Logistics", "Facturacion", "RutaExportarArchivoFacturas")
  VerRecibos
End Sub

Private Sub VerRecibos()
  Dim strSql As String
  LstRecibos.ListItems.Clear
  strSql = "SELECT recibos_caja.*, terceros.RazonSocial " & _
                          "FROM recibos_caja " & _
                          "LEFT JOIN terceros ON recibos_caja.IdTercero = terceros.IdTercero " & _
                          "WHERE Exportado=0 AND Impreso = 1 AND numero <> 0"
  rstRecibosExp.Open strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
  IniProg rstRecibosExp.RecordCount
  If rstRecibosExp.RecordCount > 0 Then
    Do While rstRecibosExp.EOF = False
      Prog (rstRecibosExp.AbsolutePosition)
      Set Item = LstRecibos.ListItems.Add(, , rstRecibosExp!IdRecibo)
      Item.SubItems(1) = rstRecibosExp.Fields("numero")
      Item.SubItems(2) = Format(rstRecibosExp!FechaPago, "dd/mm/yy")
      Item.SubItems(3) = rstRecibosExp!RazonSocial & ""
      Item.SubItems(4) = rstRecibosExp!Total & ""
      rstRecibosExp.MoveNext
    Loop
  End If
  FinProg
  rstRecibosExp.Close
End Sub
