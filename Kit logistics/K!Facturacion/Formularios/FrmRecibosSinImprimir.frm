VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRecibosSinImprimir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recibos sin imprimir"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   9840
      TabIndex        =   0
      Top             =   5760
      Width           =   1335
   End
   Begin MSComctlLib.ListView LstRecibos 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9763
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Numero"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tercero"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FrmRecibosSinImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstRecibosExp As New ADODB.Recordset
Private Sub VerRecibos()
  Dim strSql As String
  LstRecibos.ListItems.Clear
  strSql = "SELECT recibos_caja.*, terceros.RazonSocial " & _
                          "FROM recibos_caja " & _
                          "LEFT JOIN terceros ON recibos_caja.IdTercero = terceros.IdTercero " & _
                          "WHERE Impreso = 0"
  rstRecibosExp.Open strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
  IniProg rstRecibosExp.RecordCount
  If rstRecibosExp.RecordCount > 0 Then
    Do While rstRecibosExp.EOF = False
      Prog (rstRecibosExp.AbsolutePosition)
      Set Item = LstRecibos.ListItems.Add(, , rstRecibosExp!IdRecibo)
      Item.SubItems(1) = rstRecibosExp.Fields("numero")
      Item.SubItems(2) = Format(rstRecibosExp!FechaPago, "dd/mm/yy")
      Item.SubItems(3) = rstRecibosExp!RazonSocial & ""
      Item.SubItems(4) = rstRecibosExp!Total & ""
      rstRecibosExp.MoveNext
    Loop
  End If
  FinProg
  rstRecibosExp.Close
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  VerRecibos
End Sub
