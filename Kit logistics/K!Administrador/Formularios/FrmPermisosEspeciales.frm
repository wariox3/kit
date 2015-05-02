VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPermisosEspecialesUsuario 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Permisos especiales..."
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdQuitar 
      Caption         =   "<"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   ">"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin MSComctlLib.ListView LstPermisos 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   9128
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
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Permiso"
         Object.Width           =   4586
      EndProperty
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   255
      Left            =   6480
      TabIndex        =   0
      Top             =   5520
      Width           =   1815
   End
   Begin MSComctlLib.ListView LstUsuariosPermisos 
      Height          =   5175
      Left            =   4560
      TabIndex        =   4
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   9128
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
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Permiso"
         Object.Width           =   4586
      EndProperty
   End
   Begin VB.Label LblIdUsuario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label LblNmUsuario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   5400
      Width           =   3135
   End
End
Attribute VB_Name = "FrmPermisosEspecialesUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAgregar_Click()
On Error GoTo ElErr
  If LstPermisos.ListItems.Count > 0 Then
    If rstUniversal.State = adStateOpen Then rstUniversal.Close
    rstUniversal.Open "Insert into usupermisosesp (IdUsuario, IdPermiso) values (" & Val(LblIdUsuario.Caption) & ", " & LstPermisos.SelectedItem & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
    LlenarPermisos
  End If
ElErr:
  If Err.Number = -2147217900 Then
    MsgBox "El permiso ya esta asignado a este usuario", vbInformation
  End If
End Sub

Private Sub CmdAgregarTodos_Click()

End Sub

Private Sub CmdQuitar_Click()
  If LstUsuariosPermisos.ListItems.Count > 0 Then
    If rstUniversal.State = adStateOpen Then rstUniversal.Close
    rstUniversal.Open "Delete from usupermisosesp where Idusuario=" & Val(LblIdUsuario.Caption) & " and IdPermiso =" & LstUsuariosPermisos.SelectedItem, CnnPrincipal, adOpenDynamic, adLockOptimistic
    LlenarPermisos
  End If
End Sub



Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  LlenarPermisos
  LblIdUsuario.Caption = FufuLo
  LblNmUsuario.Caption = FufuSt
End Sub

Private Sub LlenarPermisos()
  LstPermisos.ListItems.Clear
  AbrirRecorset rstUniversal, "Select permisosespeciales.* from permisosespeciales", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    Set Item = LstPermisos.ListItems.Add(, , rstUniversal!IdPermiso)
      Item.SubItems(1) = rstUniversal!NmPermiso & ""
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
  
  LstUsuariosPermisos.ListItems.Clear
  AbrirRecorset rstUniversal, "select Idusuario, usupermisosesp.IdPermiso, NmPermiso From (usupermisosesp join permisosespeciales on((usupermisosesp.IdPermiso = permisosespeciales.IdPermiso))) where IdUsuario=" & FufuLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    Set Item = LstUsuariosPermisos.ListItems.Add(, , rstUniversal!IdPermiso)
      Item.SubItems(1) = rstUniversal!NmPermiso & ""
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
End Sub

