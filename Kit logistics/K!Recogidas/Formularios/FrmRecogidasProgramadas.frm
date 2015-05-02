VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRecogidasProgramadas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Recogidas Programadas"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "Eliminar marcados"
      Height          =   255
      Left            =   6240
      TabIndex        =   9
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "Agregar"
      Height          =   255
      Left            =   7800
      TabIndex        =   8
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox TxtNmCliente 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      TabIndex        =   7
      Top             =   4560
      Width           =   4695
   End
   Begin VB.TextBox TxtIdCliente 
      Height          =   285
      Left            =   3000
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox TxtHora 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Text            =   "12"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox TxtMin 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin MSComctlLib.ListView LstRecogidasProgramadas 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   7646
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cliente"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Hora"
         Object.Width           =   1587
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      Height          =   195
      Left            =   2400
      TabIndex        =   5
      Top             =   4560
      Width           =   525
   End
   Begin VB.Label Label1 
      Caption         =   "Hora:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Min:"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   4560
      Width           =   375
   End
End
Attribute VB_Name = "FrmRecogidasProgramadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LlenarRecogidas()
  LstRecogidasProgramadas.ListItems.Clear
  AbrirRecorset rstUniversal, "Select anunciosprogramados.*, terceros.RazonSocial from anunciosprogramados left join terceros on anunciosprogramados.IdCliente=terceros.IdTercero where 1 ORDER BY terceros.RazonSocial", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Do While rstUniversal.EOF = False
      Set Item = LstRecogidasProgramadas.ListItems.Add(, , rstUniversal!IdAnuncioProgramado)
      Item.SubItems(1) = rstUniversal!IdCliente & ""
      Item.SubItems(2) = rstUniversal!RazonSocial & ""
      Item.SubItems(3) = Format(rstUniversal!Hora, "hh:mm")
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
End Sub

Private Sub CmdAgregar_Click()
If TxtIdCliente.Text <> "" Then
  AbrirRecorset rstUniversal, "Insert into anunciosprogramados(IdCliente, Hora) values(" & TxtIdCliente.Text & ",'" & TxtHora.Text & ":" & TxtMin.Text & ":00')", CnnPrincipal, adOpenDynamic, adLockOptimistic
  LlenarRecogidas
  TxtIdCliente.Text = ""
  TxtNmCliente.Text = ""
End If
End Sub

Private Sub CmdEliminar_Click()
If LstRecogidasProgramadas.ListItems.Count > 0 Then
  II = 1
  Do While II <= LstRecogidasProgramadas.ListItems.Count
    If LstRecogidasProgramadas.ListItems(II).Checked = True Then
      AbrirRecorset rstUniversal, "delete from anunciosprogramados where IdAnuncioProgramado=" & Val(LstRecogidasProgramadas.SelectedItem), CnnPrincipal, adOpenDynamic, adLockOptimistic
      LstRecogidasProgramadas.ListItems.Remove II
    Else
      II = II + 1
    End If
  Loop
  LlenarRecogidas
Else
  MsgBox "No hay registro seleccionados", vbCritical
End If
End Sub

Private Sub Form_Load()
  LlenarRecogidas
End Sub

Private Sub TxtIdCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsulta 7, CnnPrincipal
    TxtIdCliente.Text = Principal.ToolConsultas1.DatSt
  End If
End Sub

Private Sub TxtIdCliente_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtIdCliente, KeyAscii, 1
End Sub

Private Sub TxtIdCliente_Validate(Cancel As Boolean)
  If Val(TxtIdCliente) <> 0 Then
    AbrirRecorset rstUniversal, "SELECT IdTercero, RazonSocial, Direccion, Telefono From Terceros Where IdTercero ='" & TxtIdCliente & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      TxtNmCliente.Text = rstUniversal!RazonSocial & ""
    Else
      TxtNmCliente.Text = "": TxtIdCliente.Text = ""
    End If
    CerrarRecorset rstUniversal
  End If
End Sub
