VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVerCuentaCobrar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalles cuenta cobrar"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   10320
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin MSComctlLib.ListView LstReciboDet 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   3836
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Numero"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Descuento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Aj. peso"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Rte ICA"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Rte Fte"
         Object.Width           =   2117
      EndProperty
   End
End
Attribute VB_Name = "FrmVerCuentaCobrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  VerDetalle
End Sub

Private Sub VerDetalle()
  Dim rstRecibosDetalle As New ADODB.Recordset
  rstRecibosDetalle.CursorLocation = adUseClient
  AbrirRecorset rstRecibosDetalle, "SELECT recibos_caja_det.*, numero, FechaPago " & _
                          "FROM recibos_caja_det " & _
                          "LEFT JOIN recibos_caja ON recibos_caja_det.IdRecibo = recibos_caja.IdRecibo " & _
                          "LEFT JOIN cuentas_cobrar ON recibos_caja_det.codigo_cuenta_cobrar_fk = cuentas_cobrar.IdCxC " & _
                          "WHERE codigo_cuenta_cobrar_fk = " & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
  LstReciboDet.ListItems.Clear
  If rstRecibosDetalle.RecordCount > 0 Then
    Do While rstRecibosDetalle.EOF = False
      Set Item = LstReciboDet.ListItems.Add(, , rstRecibosDetalle!IdReciboDet)
      Item.SubItems(1) = rstRecibosDetalle!Numero & ""
      Item.SubItems(2) = rstRecibosDetalle!FechaPago & ""
      Item.SubItems(3) = rstRecibosDetalle.Fields("valor")
      Item.SubItems(4) = rstRecibosDetalle!descuento
      Item.SubItems(5) = rstRecibosDetalle.Fields("ajuste_peso")
      Item.SubItems(6) = rstRecibosDetalle.Fields("retencion_ica")
      Item.SubItems(7) = rstRecibosDetalle.Fields("retencion_fuente")
      rstRecibosDetalle.MoveNext
    Loop
  End If
End Sub

