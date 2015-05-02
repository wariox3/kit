VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBuscarCO 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buscar Centro de Operaciones"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6165
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtCOSel 
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   5280
      Width           =   1335
   End
   Begin MSComctlLib.TreeView TreeCO 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8916
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImOpciones"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImOpciones 
      Left            =   0
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarCO.frx":0000
            Key             =   "S"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarCO.frx":059A
            Key             =   "P"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CO Seleccionado:"
      Height          =   195
      Left            =   2520
      TabIndex        =   4
      Top             =   5280
      Width           =   1290
   End
End
Attribute VB_Name = "FrmBuscarCO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCO As New ADODB.Recordset
Dim Node As Node

Private Sub CmdAceptar_Click()
  If Val(TxtCOSel.Text) <> 0 Then
    FufuLo = Val(TxtCOSel.Text)
    Unload Me
  End If
End Sub

Private Sub CmdCancelar_Click()
  FufuLo = 0
  Unload Me
End Sub

Private Sub Form_Load()
  LlenarTree
End Sub

Private Sub LlenarTree()
  TreeCO.Nodes.Clear
  rstCO.Open "Select*from CentrosOperaciones where Tipo=0", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Do While rstCO.EOF = False
      Set Node = TreeCO.Nodes.Add(, , "M" & rstCO!IdPo, rstCO!NmPuntoOperaciones & "", "P")
      rstCO.MoveNext
    Loop
  rstCO.Close
  rstCO.Open "Select*from CentrosOperaciones where Tipo<>0", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Do While rstCO.EOF = False
      Set Node = TreeCO.Nodes.Add("M" & rstCO!Tipo, tvwChild, "M" & rstCO!IdPo, rstCO!NmPuntoOperaciones & "", "S")
      rstCO.MoveNext
    Loop
  rstCO.Close
End Sub

Private Sub TreeCO_NodeClick(ByVal Node As MSComctlLib.Node)
  TxtCOSel.Text = Mid(TreeCO.Nodes(TreeCO.SelectedItem.Index).Key, 2)
End Sub
