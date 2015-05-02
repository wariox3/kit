VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVerAuxiliares 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Auxiliares de la recogida..."
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "Agregar"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton CmdQuitar 
      Caption         =   "Quitar seleccionado"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin MSComctlLib.ListView LstAuxiliares 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   2990
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Auxiliar"
         Object.Width           =   8819
      EndProperty
   End
End
Attribute VB_Name = "FrmVerAuxiliares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAgregar_Click()
  On Error GoTo ExisteRegistro
  Principal.ToolConsultas1.AbrirDevConsulta 10, CnnPrincipal
  If Principal.ToolConsultas1.DatSt <> "" Then
    AbrirRecorset rstUniversal, "INSERT INTO AuxiliaresVehiculos VALUES (" & FufuLo & ",'" & Principal.ToolConsultas1.DatSt & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
    MsgBox "Auxiliar agregado con exito", vbInformation, "Auxiliar agregado"
    VerAuxiliares
  End If
ExisteRegistro:
  If Err.Number = -2147217900 Then MsgBox "Ya existe este auxiliar para este vehiculo", vbCritical, "Ya existe el auxiliar"
End Sub

Private Sub CmdQuitar_Click()
On Error GoTo SinItem
  If MsgBox("¿Esta seguro de eliminar el auxiliar " & LstAuxiliares.ListItems(LstAuxiliares.SelectedItem.Index).SubItems(1) & "?", vbQuestion + vbYesNo) = vbYes Then
    AbrirRecorset rstUniversal, "Delete from AuxiliaresVehiculos where IdAsignacion=" & Val(Me.Tag) & " and IdAuxiliar='" & LstAuxiliares.SelectedItem & "'", CnnPrincipal, adOpenDynamic, adLockBatchOptimistic
    LstAuxiliares.ListItems.Remove LstAuxiliares.SelectedItem.Index
    MsgBox "El auxiliar se ha eliminado con exito", vbInformation
  End If
SinItem:
  If Err.Number = 91 Then MsgBox "No hay un auxiliar seleccionado para eliminar", vbCritical
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Tag = FufuLo
  VerAuxiliares
End Sub
Private Sub VerAuxiliares()
  LstAuxiliares.ListItems.Clear
  AbrirRecorset rstUniversal, "SELECT Auxiliares.NmAuxiliar, AuxiliaresVehiculos.IdAsignacion, AuxiliaresVehiculos.IdAuxiliar FROM AuxiliaresVehiculos INNER JOIN Auxiliares ON AuxiliaresVehiculos.IdAuxiliar = Auxiliares.IdAuxiliar where IdAsignacion=" & Val(Me.Tag), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    Set Item = LstAuxiliares.ListItems.Add(, , rstUniversal!IdAuxiliar)
     Item.SubItems(1) = rstUniversal!NmAuxiliar & ""
     rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
End Sub
