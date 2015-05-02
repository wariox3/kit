VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmControlFacturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control facturas..."
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdMarcarComoNoExportadas 
      Caption         =   "Marcar seleccionadas como no exportada"
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   6600
      Width           =   3975
   End
   Begin VB.CheckBox ChkSoloNoExportadas 
      Caption         =   "Solo no exportadas"
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   6960
      Width           =   1935
   End
   Begin VB.ComboBox CboTipo 
      Height          =   315
      ItemData        =   "FrmControlFacturas.frx":0000
      Left            =   720
      List            =   "FrmControlFacturas.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   7320
      Width           =   2055
   End
   Begin VB.CheckBox ChkSoloExportadas 
      Caption         =   "Solo exportadas"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9240
      TabIndex        =   6
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   6600
      Width           =   1815
   End
   Begin MSComctlLib.ListView LstFacturas 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   11245
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
         Text            =   "Factura"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tp"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tercero"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Exp"
         Object.Width           =   882
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPFechaDesde 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   6600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yy"
      Format          =   16777219
      CurrentDate     =   39740
   End
   Begin MSComCtl2.DTPicker DTPFechaHasta 
      Height          =   315
      Left            =   720
      TabIndex        =   2
      Top             =   6960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yy"
      Format          =   16777219
      CurrentDate     =   39740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   7320
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   6960
      Width           =   465
   End
   Begin VB.Label LblFecha 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   6600
      Width           =   495
   End
End
Attribute VB_Name = "FrmControlFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstFacturasExp As New ADODB.Recordset

Private Sub CmdBuscar_Click()
  VerFacturas
End Sub

Private Sub CmdMarcarComoNoExportadas_Click()
  Dim rstFactura As New ADODB.Recordset
  rstFactura.CursorLocation = adUseClient
  Dim strSql As String
  II = 1
  While II <= LstFacturas.ListItems.Count
    If LstFacturas.ListItems(II).Checked = True Then
      AbrirRecorset rstUniversal, "SELECT Exportada FROM facturas_venta WHERE TipoFactura = " & LstFacturas.ListItems(II).SubItems(1) & " AND Numero =" & LstFacturas.ListItems(II) & " AND Exportada = 1", CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.RecordCount > 0 Then
        rstFactura.Open "UPDATE facturas_venta SET Exportada=0 where Numero=" & LstFacturas.ListItems(II) & " AND TipoFactura = " & LstFacturas.ListItems(II).SubItems(1), CnnPrincipal, adOpenDynamic, adLockOptimistic
      End If
      LstFacturas.ListItems.Remove (II)
    Else
     II = II + 1
    End If
  Wend
VerFacturas
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  rstFacturasExp.CursorLocation = adUseClient
  DTPFechaDesde.Value = Date
  DTPFechaHasta.Value = Date
End Sub
Private Sub VerFacturas()
  Dim strSql As String
  strSql = "SELECT facturas_venta.*, terceros.RazonSocial, NmTipoFactura " & _
                          "FROM facturas_venta " & _
                          "LEFT JOIN terceros ON facturas_venta.IdTercero = terceros.IdTercero " & _
                          "LEFT JOIN facturas_tipos ON facturas_venta.TipoFactura = facturas_tipos.IdTipoFactura " & _
                          "WHERE Fecha >='" & Format(DTPFechaDesde.Value, "yyyy/mm/dd") & "' AND Fecha <='" & Format(DTPFechaHasta.Value, "yyyy/mm/dd") & "'"
  If ChkSoloExportadas.Value = 1 Then
    strSql = strSql & " AND Exportada = 1"
  End If
  If ChkSoloNoExportadas.Value = 1 Then
    strSql = strSql & " AND Exportada = 0"
  End If
  If CboTipo.ListIndex >= 1 Then
    strSql = strSql & " AND TipoFactura = " & CboTipo.ListIndex
  End If
  strSql = strSql & " ORDER BY Numero"
  LstFacturas.ListItems.Clear
  AbrirRecorset rstFacturasExp, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
  IniProg rstFacturasExp.RecordCount
  If rstFacturasExp.RecordCount > 0 Then
    Do While rstFacturasExp.EOF = False
      Prog (rstFacturasExp.AbsolutePosition)
      Set Item = LstFacturas.ListItems.Add(, , rstFacturasExp!Numero)
      Item.SubItems(1) = rstFacturasExp!TipoFactura
      Item.SubItems(2) = rstFacturasExp!NmTipoFactura
      Item.SubItems(3) = Format(rstFacturasExp!Fecha, "dd/mm/yy")
      Item.SubItems(4) = rstFacturasExp!RazonSocial & ""
      If Val(rstFacturasExp!Exportada) = 1 Then
        Item.SubItems(5) = "SI"
      Else
        Item.SubItems(5) = "NO"
      End If
      rstFacturasExp.MoveNext
    Loop
  End If
  FinProg
  rstFacturasExp.Close
End Sub

