VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVerRegistroEnvioEmail 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro envio emails novedades"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   255
      Left            =   8520
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
   End
   Begin MSComctlLib.ListView LstRegistros 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha Envio"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Mail Envio"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Mail Destino"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Usuario"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "FrmVerRegistroEnvioEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim rstRegistros As New ADODB.Recordset
  rstRegistros.CursorLocation = adUseClient
  FufuSt = "SELECT registro_envios_email.*, usuarios.NmUsuario " & _
           "FROM registro_envios_email " & _
           "LEFT JOIN usuarios ON registro_envios_email.Usuario = usuarios.IDUsuario " & _
           "WHERE Guia = " & FufuLo
  AbrirRecorset rstRegistros, FufuSt, CnnPrincipal, adOpenDynamic, adLockOptimistic
  Do While rstRegistros.EOF = False
    Set Item = LstRegistros.ListItems.Add(, , rstRegistros!FechaEnvio)
    Item.SubItems(1) = rstRegistros!MailEnvio & ""
    Item.SubItems(2) = rstRegistros!MailDestino & ""
    Item.SubItems(3) = rstRegistros!NmUsuario & ""
    rstRegistros.MoveNext
  Loop
  CerrarRecorset rstRegistros
  
  Set rstRegistros = Nothing
End Sub

