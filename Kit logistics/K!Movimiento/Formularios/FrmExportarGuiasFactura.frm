VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmExportarGuiasFactura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar guias factura"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   13215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExportarPorGuia 
      Caption         =   "Exportar por guia"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Orden"
      Height          =   855
      Left            =   6360
      TabIndex        =   8
      Top             =   5160
      Width           =   2295
      Begin VB.OptionButton Option2 
         Caption         =   "Tipo cobro"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton OptOrdFecha 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.CommandButton CmdActualizar 
      Caption         =   "Actualizar / buscar"
      Height          =   375
      Left            =   8760
      TabIndex        =   7
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   10680
      TabIndex        =   2
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton CmdExportarSeleccionadas 
      Caption         =   "Exportar seleccionadas"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   2415
   End
   Begin MSComctlLib.ListView LstGuias 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   8916
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Guia"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Documento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cliente"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Destino"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Flete"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Manejo"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Total"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Cobro"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPFechaDesde 
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   5280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yy"
      Format          =   16777219
      CurrentDate     =   39740
   End
   Begin MSComCtl2.DTPicker DTPFechaHasta 
      Height          =   315
      Left            =   4560
      TabIndex        =   4
      Top             =   5640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yy"
      Format          =   16777219
      CurrentDate     =   39740
   End
   Begin VB.Label LblFecha 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   195
      Left            =   3960
      TabIndex        =   6
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   4080
      TabIndex        =   5
      Top             =   5640
      Width           =   465
   End
End
Attribute VB_Name = "FrmExportarGuiasFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstGuiasFactura As New ADODB.Recordset

Private Sub CmdActualizar_Click()
  verGuiasFactura
End Sub

Private Sub CmdExportarPorGuia_Click()
  If Principal.ToolConsultas1.AbrirDevDatos("Digite el numero de la guia", "Digite el numero de la guia que desea entregar", 3, 0) = True Then
    AbrirRecorset rstUniversal, "Select Guia from guias where ExportadaCartera = 0 AND Guia=" & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstUniversal.RecordCount > 0 Then
      ExportarGuiaFactura Principal.ToolConsultas1.DatLo
      CmdExportarPorGuia_Click
    Else
      MsgBox "No se encontro la guia o ya estaba exportada", vbCritical
      CmdExportarPorGuia_Click
    End If
    CerrarRecorset rstUniversal
  End If
End Sub

Private Sub CmdExportarSeleccionadas_Click()
  II = 1
  While II <= LstGuias.ListItems.Count
    If LstGuias.ListItems(II).Checked = True Then
      ExportarGuiaFactura LstGuias.ListItems(II)
      LstGuias.ListItems.Remove (II)
    Else
     II = II + 1
    End If
  Wend
End Sub



Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  rstGuiasFactura.CursorLocation = adUseClient
  DTPFechaDesde.value = Date
  DTPFechaHasta.value = Date
  verGuiasFactura
End Sub

Private Sub verGuiasFactura()

           
  LstGuias.ListItems.Clear
  AbrirRecorset rstGuiasFactura, DevSql(), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstGuiasFactura.RecordCount > 0 Then
    Do While rstGuiasFactura.EOF = False
      Set Item = LstGuias.ListItems.Add(, , rstGuiasFactura!Guia)
      Item.SubItems(1) = Format(rstGuiasFactura!FhEntradaBodega, "yyyy/mm/dd")
      Item.SubItems(2) = rstGuiasFactura!DocCliente
      Item.SubItems(3) = rstGuiasFactura!Cliente
      Item.SubItems(4) = rstGuiasFactura!NmCiudad
      Item.SubItems(5) = rstGuiasFactura!VrFlete
      Item.SubItems(6) = rstGuiasFactura!VrManejo
      Item.SubItems(7) = rstGuiasFactura!VrManejo + rstGuiasFactura!VrFlete
      Item.SubItems(8) = rstGuiasFactura!NmTipoCobro
      rstGuiasFactura.MoveNext
    Loop
  End If
  CerrarRecorset rstGuiasFactura
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstGuiasFactura = Nothing
End Sub

Private Function DevSql() As String
  Dim strSql As String
  strSql = "SELECT Guia, FhEntradaBodega, DocCliente, Cliente, NmCiudad, VrFlete, VrManejo, NmTipoCobro " & _
           "FROM guias " & _
           "LEFT JOIN ciudades ON guias.IdCiuDestino = ciudades.IdCiudad " & _
           "LEFT JOIN tipos_cobro ON guias.TipoCobro = tipos_cobro.IdTipoCobro " & _
           "WHERE ExportadaCartera = 0 AND (TipoCobro = 1 OR TipoCobro = 2) AND FhEntradaBodega >='" & Format(DTPFechaDesde.value, "yyyy/mm/dd") & " 00:00:00' AND FhEntradaBodega <='" & Format(DTPFechaHasta.value, "yyyy/mm/dd") & " 23:59:00' "
           
  If OptOrdFecha.value = True Then
    strSql = strSql & " ORDER BY FhEntradaBodega ASC"
  Else
    strSql = strSql & " ORDER BY TipoCobro ASC"
  End If
  DevSql = strSql
    
End Function
