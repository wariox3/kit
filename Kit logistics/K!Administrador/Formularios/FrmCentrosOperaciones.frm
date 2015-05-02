VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCentrosOperaciones 
   Caption         =   "Centros de Operaciones"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   6240
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtCOSel 
      Height          =   285
      Left            =   5160
      TabIndex        =   6
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton CmdEditarCO 
      Caption         =   "Editar CO"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton CmdNuevoAcopio 
      Caption         =   "Nuevo Acopio"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "Nuevo CO"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton CmdEliminarCO 
      Caption         =   "Eliminar CO"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "Cerrar"
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Top             =   5280
      Width           =   855
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
      Left            =   4200
      Top             =   4080
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
            Picture         =   "FrmCentrosOperaciones.frx":0000
            Key             =   "S"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentrosOperaciones.frx":059A
            Key             =   "P"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Centro de operaciones seleccionado:"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   5640
      Width           =   2775
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu MnuNewCO 
         Caption         =   "Nuevo Centro de Operaciones"
      End
      Begin VB.Menu MnuNewAcopio 
         Caption         =   "Nuevo Acopio"
      End
      Begin VB.Menu MnuEditCO 
         Caption         =   "Editar Centro de Operaciones"
      End
      Begin VB.Menu MnuEliminarCO 
         Caption         =   "Eliminar Centro de Operaciones"
      End
   End
End
Attribute VB_Name = "FrmCentrosOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Node As Node

Private Sub CmdCerrar_Click()
  FormAbierto = False
  Unload Me
End Sub

Private Sub CmdEditarCO_Click()
    AbrirRecorset rstUniversal, "Select IdPo, IdCiudad, NmPuntoOperaciones from CentrosOperaciones where IdPo=" & Val(Mid(TreeCO.Nodes(TreeCO.SelectedItem.Index).Key, 2)), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        II = rstUniversal!IdCiudad
        FufuSt = rstUniversal!NmPuntoOperaciones & ""
      End If
    CerrarRecorset rstUniversal
    AbrirRecorset rstUniversal, "Select IdCiudad, NmCiudad from ciudades where IdCiudad=" & II, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        FrmAddCO.CboCiudades.Tag = II
        FrmAddCO.CboCiudades.Text = rstUniversal!NmCiudad & ""
      End If
    CerrarRecorset rstUniversal
    II = 3
    FrmAddCO.TxtCO = Mid(TreeCO.Nodes(TreeCO.SelectedItem.Index).Key, 2)
    FrmAddCO.TxtNombreCO = FufuSt
    FrmAddCO.Show 1
    LlenarTree
End Sub

Private Sub CmdEliminarCO_Click()
  If TreeCO.Nodes(TreeCO.SelectedItem.Index).Image = "P" Then
    If MsgBox("Va a eliminar un centro de operaciones principal, si elimina este elemento, automaticamente eliminara todos los Acopios que dependen de este" & Chr(13) & "¿Desea elimnar el CO?", vbQuestion + vbYesNo) = vbYes Then
      AbrirRecorset rstUniversal, "Delete from CentrosOperaciones where IdPo=" & Mid(TreeCO.Nodes(TreeCO.SelectedItem.Index).Key, 2), CnnPrincipal, adOpenDynamic, adLockOptimistic
      AbrirRecorset rstUniversal, "Delete from CentrosOperaciones where Tipo=" & Mid(TreeCO.Nodes(TreeCO.SelectedItem.Index).Key, 2), CnnPrincipal, adOpenDynamic, adLockOptimistic
      MsgBox "El centro de operaciones y todos sus acopios ha sido eliminado en exito", vbInformation
    End If
  Else
    If MsgBox("va a eliminar el acopio [" & TreeCO.Nodes(TreeCO.SelectedItem.Index) & "]" & Chr(13) & "¿Esta seguro de eliminar este acopio?", vbQuestion + vbYesNo) = vbYes Then
      AbrirRecorset rstUniversal, "Delete from CentrosOperaciones where IdPo=" & Mid(TreeCO.Nodes(TreeCO.SelectedItem.Index).Key, 2), CnnPrincipal, adOpenDynamic, adLockOptimistic
      MsgBox "El acopio ha sido eliminado con exito", vbInformation
    End If
  End If
  LlenarTree
End Sub

Private Sub CmdNuevo_Click()
  II = 1
  FrmAddCO.Show 1
  LlenarTree
End Sub

Private Sub CmdNuevoAcopio_Click()
  If TreeCO.Nodes(TreeCO.SelectedItem.Index).Image = "P" Then
    II = 2
    FrmAddCO.TxtCO = Mid(TreeCO.Nodes(TreeCO.SelectedItem.Index).Key, 2)
    FrmAddCO.Show 1
    LlenarTree
  Else
    MsgBox "Debe seleccionar un centro de operaciones principal", vbCritical
  End If
End Sub
Private Sub Form_Load()
  LlenarTree
End Sub
Private Sub LlenarTree()
  TreeCO.Nodes.Clear
  AbrirRecorset rstUniversal, "Select centrosoperaciones.* from centrosoperaciones where Tipo=0", CnnPrincipal, adOpenDynamic, adLockOptimistic
    Do While rstUniversal.EOF = False
      Set Node = TreeCO.Nodes.Add(, , "M" & rstUniversal!IdPo, rstUniversal!NmPuntoOperaciones & "", "P")
      rstUniversal.MoveNext
    Loop
  CerrarRecorset rstUniversal
  AbrirRecorset rstUniversal, "Select*from CentrosOperaciones where Tipo<>0", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Do While rstUniversal.EOF = False
      Set Node = TreeCO.Nodes.Add("M" & rstUniversal!Tipo, tvwChild, "M" & rstUniversal!IdPo, rstUniversal!NmPuntoOperaciones & "", "S")
      rstUniversal.MoveNext
    Loop
  CerrarRecorset rstUniversal
End Sub

Private Sub MnuEditCO_Click()
  CmdEditarCO_Click
End Sub

Private Sub MnuEliminarCO_Click()
  CmdEliminarCO_Click
End Sub

Private Sub MnuNewAcopio_Click()
  CmdNuevoAcopio_Click
End Sub

Private Sub MnuNewCO_Click()
  CmdNuevo_Click
End Sub

Private Sub TreeCO_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
    PopupMenu MnuArchivo
  End If
End Sub

Private Sub TreeCO_NodeClick(ByVal Node As MSComctlLib.Node)
  TxtCOSel.Text = Mid(TreeCO.Nodes(TreeCO.SelectedItem.Index).Key, 2)
End Sub
