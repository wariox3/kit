VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBuscarNegociaciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccionar negociacion..."
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   4080
      Width           =   1695
   End
   Begin MSComctlLib.ListView LstNegociaciones 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6800
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Negociacion"
         Object.Width           =   8819
      EndProperty
   End
End
Attribute VB_Name = "FrmBuscarNegociaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
If LstNegociaciones.ListItems.Count > 0 Then
  FufuLo = LstNegociaciones.SelectedItem
  Unload Me
Else
  MsgBox "No hay una negociacion seleccionada", vbCritical
End If
End Sub

Private Sub CmdCancelar_Click()
  FufuLo = 0
  Unload Me
End Sub

Private Sub VerNegociaciones()
  LstNegociaciones.ListItems.Clear
  AbrirRecorset rstUniversal, "SELECT negociaciones_terceros.IdTercero,  negociaciones_terceros.IdNegociacion, negociaciones_terceros.Activo, negociaciones.NmNegociacion From negociaciones_terceros INNER JOIN negociaciones ON (negociaciones_terceros.IdNegociacion = negociaciones.Id) where IdTercero='" & FufuSt & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    Set Item = LstNegociaciones.ListItems.Add(, , rstUniversal.Fields("IdNegociacion") & "")
        Item.SubItems(1) = rstUniversal.Fields("NmNegociacion")
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
End Sub

Private Sub Form_Load()
  VerNegociaciones
End Sub

Private Sub LstNegociaciones_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
