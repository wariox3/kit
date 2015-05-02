VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAgregarFormulario 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar formulario"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LstFormularios 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5953
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Formulario"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar / Agregar"
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   3600
      Width           =   1575
   End
End
Attribute VB_Name = "FrmAgregarFormulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstNuevoFormulario As New ADODB.Recordset

Private Sub CmdAceptar_Click()
  Dim II As Integer
  II = 1
  While II <= LstFormularios.ListItems.Count
    If LstFormularios.ListItems(II).Checked = True Then
      AbrirRecorset rstUniversal, "Select*from permisos where IdUsuario=" & FufuLo & " and IdFormulario=" & LstFormularios.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.RecordCount <= 0 Then
        If rstNuevoFormulario.State = adStateOpen Then rstNuevoFormulario.Close
        rstNuevoFormulario.Open "Insert into permisos (IdUsuario, IdFormulario, Ingreso, Nuevo, Editar, Eliminar) values (" & FufuLo & ", " & LstFormularios.ListItems(II) & ", 1, 1, 1, 1)", CnnPrincipal, adOpenDynamic, adLockOptimistic
      End If
      CerrarRecorset rstUniversal
    End If
    II = II + 1
  Wend
  Unload Me
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  rstNuevoFormulario.CursorLocation = adUseClient
  LlenarFormularios
End Sub
Private Sub LlenarFormularios()
  LstFormularios.ListItems.Clear
  AbrirRecorset rstUniversal, "Select * from Formularios", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    Set Item = LstFormularios.ListItems.Add(, , rstUniversal!IdFormulario)
      Item.SubItems(1) = rstUniversal!NmFormulario & ""
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
End Sub
