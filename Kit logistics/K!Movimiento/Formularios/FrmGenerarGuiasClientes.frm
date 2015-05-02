VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmGenerarGuiasClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generar guias clientes"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   5760
      Width           =   1815
   End
   Begin MSComctlLib.ListView LstLotesGuias 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nit"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tercero"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Desde"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Hasta"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FrmGenerarGuiasClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub VerLotes()
  LstLotesGuias.ListItems.Clear
  AbrirRecorset rstUniversal, "SELECT IdLote, guias_lotes.IdTercero, RazonSocial, Desde, Hasta FROM guias_lotes LEFT JOIN terceros ON guias_lotes.IdTercero = terceros.IDTercero WHERE 1 ORDER BY IdLote ASC", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Do While rstUniversal.EOF = False
      Set Item = LstLotesGuias.ListItems.Add(, , rstUniversal!IdLote)
      Item.SubItems(1) = rstUniversal!IdTercero
      Item.SubItems(2) = rstUniversal!RazonSocial
      Item.SubItems(3) = rstUniversal!Desde
      Item.SubItems(4) = rstUniversal!Hasta
      rstUniversal.MoveNext
    Loop
  CerrarRecorset rstUniversal
End Sub

Private Sub CmdGenerar_Click()
  Dim douDesde As Double
  Dim douHasta As Double
  
  douDesde = LstLotesGuias.ListItems(LstLotesGuias.SelectedItem.Index).SubItems(3)
  douHasta = LstLotesGuias.ListItems(LstLotesGuias.SelectedItem.Index).SubItems(4)
  AbrirRecorset rstUniversal, "TRUNCATE TABLE guias_cliente", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Do While douDesde <= douHasta
    AbrirRecorset rstUniversal, "INSERT INTO guias_cliente VALUES (" & douDesde & ", '" & LstLotesGuias.ListItems(LstLotesGuias.SelectedItem.Index).SubItems(1) & "', '*" & douDesde & "*')", CnnPrincipal, adOpenDynamic, adLockOptimistic
    douDesde = douDesde + 1
  Loop
  Mostrar_Reporte CnnPrincipal, 41, "Select*from sql_im_impguia_clientes", "Formatos guias", 2
End Sub

Private Sub Form_Load()
  VerLotes
End Sub
