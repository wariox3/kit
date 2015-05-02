VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVerProductos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ver Productos..."
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView LstTem 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Lote"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Producto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ID"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Empaque"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Ancho"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Largo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Alto"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Cant"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "K Real"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "K Vol"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Kilos Fac"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Vr Flete"
         Object.Width           =   1940
      EndProperty
   End
End
Attribute VB_Name = "FrmVerProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
  Me.Caption = "Detalle de (Productos/Mercancia) de la guia [" & FufuLo & "]"
  LstTem.ListItems.Clear
  AbrirRecorset rstUniversal, "SELECT*From MvtoGuias Where Guia=" & FufuLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  IniProg 1, rstUniversal.RecordCount
    For II = 1 To rstUniversal.RecordCount
    Prog (II)
    Set Item = LstTem.ListItems.Add(, , rstUniversal.Fields("lote") & "")
      Item.SubItems(1) = rstUniversal.Fields("idproducto")
        AbrirRecorset rstUniversalSer, "Select IdProducto, NmProducto from Productos Where IdProducto=" & rstUniversal!IdProducto, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
          If rstUniversalSer.EOF = False Then Item.SubItems(2) = rstUniversalSer!NmProducto
        CerrarRecorset rstUniversalSer
      Item.SubItems(3) = rstUniversal.Fields("idempaque")
        AbrirRecorset rstUniversalSer, "Select IdEmpaque, NmEmpaque from Empaques where IdEmpaque=" & rstUniversal!IdEmpaque, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
          If rstUniversalSer.EOF = False Then Item.SubItems(4) = rstUniversalSer!NmEmpaque
        CerrarRecorset rstUniversalSer
      Item.SubItems(5) = rstUniversal.Fields("ancho")
      Item.SubItems(6) = rstUniversal.Fields("largo")
      Item.SubItems(7) = rstUniversal.Fields("altura")
      Item.SubItems(8) = rstUniversal.Fields("cant")
      Item.SubItems(9) = rstUniversal.Fields("KilosReal")
      Item.SubItems(10) = rstUniversal.Fields("KilosVol")
      Item.SubItems(11) = rstUniversal.Fields("KilosFacturados")
      Item.SubItems(12) = rstUniversal.Fields("VlrFlete")
      rstUniversal.MoveNext
  Next
  CerrarRecorset rstUniversal
  FinProg
End Sub
