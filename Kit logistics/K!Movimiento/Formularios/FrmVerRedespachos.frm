VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVerRedespachos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ver historial de redespachos"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdInfoDespacho 
      Caption         =   "Info Despacho"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Default         =   -1  'True
      Height          =   255
      Left            =   5880
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin MSComctlLib.ListView LstRedespachos 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Guia"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Hora"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Usuario"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Despacho"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label LblMensaje 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   4575
   End
End
Attribute VB_Name = "FrmVerRedespachos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdInfoDespacho_Click()
  If LstRedespachos.ListItems.Count > 0 Then
    FufuLo = Val(LstRedespachos.SelectedItem.SubItems(4))
    FrmInfoDespacho.Show 1
  End If
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  AbrirRecorset rstUniversal, "SELECT redespachos.*, usuarios.NmUsuario From redespachos INNER JOIN usuarios ON (redespachos.IdUsuario = usuarios.IDUsuario) where Guia=" & FufuLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = True Then LblMensaje.Caption = "Esta guia no tiene redespachos" Else LblMensaje.Caption = "Esta guia tiene " & rstUniversal.RecordCount & " redespachos"
  Do While rstUniversal.EOF = False
    Set Item = LstRedespachos.ListItems.Add(, , rstUniversal!Guia)
      Item.SubItems(1) = Format(rstUniversal!Fecha, "hh:mm:ss")
      Item.SubItems(2) = Format(rstUniversal!Fecha, "dd/mm/yy")
      Item.SubItems(3) = rstUniversal!NmUsuario & ""
      Item.SubItems(4) = rstUniversal!IdDespacho
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
End Sub

