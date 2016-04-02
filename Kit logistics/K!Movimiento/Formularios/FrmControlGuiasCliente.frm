VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmControlGuiasCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control guias cliente"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "Eliminar"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton CmdValidar 
      Caption         =   "Validar"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame FraNuevo 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   6135
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   4560
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox TxtHasta 
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TxtDesde 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxtIdTercero 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   510
      End
      Begin VB.Label LblNmCliente 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   9000
      TabIndex        =   0
      Top             =   3120
      Width           =   1695
   End
   Begin MSComctlLib.ListView LstGuiasCliente 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5106
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nit"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cliente"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Desde"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Hasta"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Estado"
         Object.Width           =   1058
      EndProperty
   End
End
Attribute VB_Name = "FrmControlGuiasCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAgregar_Click()
  Dim strSql As String
  strSql = "INSERT INTO guias_cliente (Fecha, IdTercero, Desde, Hasta, Estado) VALUES ('" & Format(Date, "yyyy/m/d h:m") & "','" & TxtIdTercero.Text & "', " & Val(TxtDesde) & ", " & Val(TxtHasta) & ", 'P')"
  AbrirRecorset rstUniversal, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
  Ver
End Sub

Private Sub CmdEliminar_Click()
  Dim rstGuiasCliente As New ADODB.Recordset
  rstGuiasCliente.CursorLocation = adUseClient
  II = 1
  While II <= LstGuiasCliente.ListItems.Count
    If LstGuiasCliente.ListItems(II).Checked = True Then
      AbrirRecorset rstGuiasCliente, "SELECT guias_cliente.* FROM guias_cliente WHERE Estado = 'P' AND IdGuiaCliente = " & LstGuiasCliente.SelectedItem, CnnPrincipal, adOpenDynamic, adLockOptimistic
        If rstGuiasCliente.RecordCount > 0 Then
          AbrirRecorset rstUniversal, "DELETE FROM guias_cliente WHERE IdGuiaCliente=" & Val(LstGuiasCliente.SelectedItem), CnnPrincipal, adOpenDynamic, adLockOptimistic
        Else
          MsgBox "No existe el rango de guias o no esta pendiente", vbCritical
        End If
      CerrarRecorset rstGuiasCliente
      LstGuiasCliente.ListItems.Remove (II)
    Else
     II = II + 1
    End If
  Wend
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Ver()
  LstGuiasCliente.ListItems.Clear
  AbrirRecorset rstUniversal, "SELECT guias_cliente.*, RazonSocial from guias_cliente left join terceros on guias_cliente.IdTercero = terceros.IDTercero where 1", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Do While rstUniversal.EOF = False
      Set Item = LstGuiasCliente.ListItems.Add(, , rstUniversal.Fields("IdGuiaCliente"))
      Item.SubItems(1) = rstUniversal.Fields("Fecha")
      Item.SubItems(2) = rstUniversal.Fields("IDTercero")
      Item.SubItems(3) = rstUniversal.Fields("RazonSocial")
      Item.SubItems(4) = rstUniversal.Fields("Desde")
      Item.SubItems(5) = rstUniversal.Fields("Hasta")
      Item.SubItems(6) = rstUniversal.Fields("Estado")
      rstUniversal.MoveNext
    Loop
  CerrarRecorset rstUniversal
End Sub

Private Sub CmdValidar_Click()
  Dim rstGuiasCliente As New ADODB.Recordset
  rstGuiasCliente.CursorLocation = adUseClient
      AbrirRecorset rstGuiasCliente, "SELECT guias_cliente.* FROM guias_cliente WHERE Estado = 'P' AND IdGuiaCliente = " & LstGuiasCliente.SelectedItem, CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstGuiasCliente.RecordCount > 0 Then
        FufuLo = rstGuiasCliente!Desde
        FufuLo2 = rstGuiasCliente!Hasta
        FrmValidarRangoGuias.Show 1
      End If
      CerrarRecorset rstGuiasCliente
End Sub

Private Sub Form_Load()
  Ver
End Sub
