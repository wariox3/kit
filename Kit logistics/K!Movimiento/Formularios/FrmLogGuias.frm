VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLogGuias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LstLog 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6588
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Accion"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Usuario"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "FrmLogGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  VerLogs
End Sub

Sub VerLogs()
  LstLog.ListItems.Clear
  rstUniversal.Open "SELECT log_guias.*, usuarios.NmUsuario, log_guias_acciones.NmAccion " & _
                    "FROM log_guias " & _
                    "LEFT JOIN usuarios ON log_guias.IdUsuario = usuarios.IDUsuario " & _
                    "LEFT JOIN log_guias_acciones ON log_guias.IdAccionLog = log_guias_acciones.IdAccionLog " & _
                    "WHERE Guia = " & FufuLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    Set Item = LstLog.ListItems.Add(, , Format(rstUniversal!Fecha, "dd/mmm/yy hh:mm:ss"))
    Item.SubItems(1) = rstUniversal!NmAccion & ""
    Item.SubItems(2) = rstUniversal!NmUsuario & ""
    rstUniversal.MoveNext
  Loop
  rstUniversal.Close
End Sub
