VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmUsuarios 
   Caption         =   "Usuarios..."
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   6210
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdPermisosFormularios 
      Caption         =   "Permisos de formularios"
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton CmdPermisosEspeciales 
      Caption         =   "Permisos especiales"
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "Eliminar"
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "Editar"
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "Nuevo"
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar >>"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton CmdActualizar 
      Caption         =   "Actualizar >>"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton CmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin MSComctlLib.ListView LstUsuarios 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8070
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
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Usuario"
         Object.Width           =   8819
      EndProperty
   End
   Begin MSComctlLib.ListView LstPermisos 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Ingresar"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Crear"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Editar"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Eliminar"
         Object.Width           =   1499
      EndProperty
   End
   Begin VB.Frame FraDatos 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton CmdCambiar 
         Caption         =   "Cambiar"
         Height          =   255
         Left            =   4680
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox ChkEliminar 
         Caption         =   "Eliminar"
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox ChkEditar 
         Caption         =   "Editar"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox ChkCrear 
         Caption         =   "Crear"
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox ChkIngreso 
         Caption         =   "Ingreso"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Menu MnuPermisos 
      Caption         =   "Permisos"
      Visible         =   0   'False
      Begin VB.Menu MnuVerPerMovimiento 
         Caption         =   "Ver permisos [Movimiento]"
      End
      Begin VB.Menu MnuVerPerVehiculos 
         Caption         =   "Ver permisos [Vehiculos]"
      End
      Begin VB.Menu MnuPermisosDatosBasicos 
         Caption         =   "Ver permisos [Datos Basicos]"
      End
      Begin VB.Menu MnuVerPermisosFacturacion 
         Caption         =   "Ver permisos [Facturacion]"
      End
      Begin VB.Menu MnuVerPermisosRecogidas 
         Caption         =   "Ver permisos [Recogidas]"
      End
   End
End
Attribute VB_Name = "FrmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdActualizar_Click()
  For II = 1 To LstPermisos.ListItems.Count
    AbrirRecorset rstUniversal, "Update permisos" & Me.Tag & " set Ingreso=" & Dev10(LstPermisos.ListItems(II).SubItems(2)) & ", Nuevo=" & Dev10(LstPermisos.ListItems(II).SubItems(3)) & ", Editar= " & Dev10(LstPermisos.ListItems(II).SubItems(4)) & ", Eliminar=" & Dev10(LstPermisos.ListItems(II).SubItems(5)) & " where IdUsuario=" & LstUsuarios.ListItems(LstUsuarios.SelectedItem.Index) & " and Formulario=" & II, CnnPrincipal, adOpenDynamic, adLockOptimistic
  Next
  LstPermisos.Visible = False
  LstUsuarios.Height = 5175
  FraDatos.Visible = False
  LstUsuarios.Enabled = True
  DesBloquearBotones
  MsgBox "Permisos actualizados con exito", vbExclamation
End Sub

Private Sub CmdCambiar_Click()
On Error GoTo SinNada
  LstPermisos.ListItems.Item(LstPermisos.SelectedItem.Index).SubItems(2) = DevSINO(ChkIngreso.Value)
  LstPermisos.ListItems.Item(LstPermisos.SelectedItem.Index).SubItems(3) = DevSINO(ChkCrear.Value)
  LstPermisos.ListItems.Item(LstPermisos.SelectedItem.Index).SubItems(4) = DevSINO(ChkEditar.Value)
  LstPermisos.ListItems.Item(LstPermisos.SelectedItem.Index).SubItems(5) = DevSINO(ChkEliminar.Value)
SinNada:
  If Err.Number = 91 Then MsgBox "No hay permisos cargados", vbCritical
End Sub

Private Sub CmdCancelar_Click()
  DesBloquearBotones
End Sub

Private Sub CmdCerrar_Click()
  FormAbierto = False
  Unload Me
End Sub
Private Sub CargarPermisos(Modulo As Byte)
  Me.Tag = Modulo
  LstPermisos.ListItems.Clear
  AbrirRecorset rstUniversal, "Select*from Permisos where IdUsuario=" & LstUsuarios.SelectedItem & " order by Formulario", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Do While rstUniversal.EOF = False
      Set Item = LstPermisos.ListItems.Add(, , rstUniversal!Formulario)
        Item.SubItems(1) = rstUniversal!NmFormulario
        Item.SubItems(2) = DevSINO(rstUniversal!Ingreso)
        Item.SubItems(3) = DevSINO(rstUniversal!Nuevo)
        Item.SubItems(4) = DevSINO(rstUniversal!Editar)
        Item.SubItems(5) = DevSINO(rstUniversal!Eliminar)
      rstUniversal.MoveNext
    Loop
  CerrarRecorset rstUniversal
  BloquearBotones
End Sub
Private Sub BloquearBotones()
  CmdCancelar.Enabled = True
  CmdActualizar.Enabled = True
  CmdEditar.Enabled = False
  CmdEliminar.Enabled = False
  CmdNuevo.Enabled = False
  LstPermisos.Visible = True
  LstUsuarios.Height = 1335
  FraDatos.Visible = True
  LstUsuarios.Enabled = False
End Sub
Private Sub DesBloquearBotones()
  CmdCancelar.Enabled = False
  CmdActualizar.Enabled = False
  CmdEditar.Enabled = True
  CmdEliminar.Enabled = True
  CmdNuevo.Enabled = True
  LstPermisos.Visible = False
  LstUsuarios.Height = 5175
  FraDatos.Visible = False
  LstUsuarios.Enabled = True
End Sub

Private Sub CmdEditar_Click()
  II = 2
  FufuLo = LstUsuarios.ListItems(LstUsuarios.SelectedItem.Index)
  FrmNuevoUsuario.Show 1
  LlenarLista
End Sub

Private Sub CmdEliminar_Click()
  If MsgBox("¿Esta seguro de eliminar este usuario?", vbQuestion + vbYesNo) = vbYes Then
    AbrirRecorset rstUniversal, "Delete from Usuarios where IdUsuario=" & LstUsuarios.ListItems(LstUsuarios.SelectedItem.Index), CnnPrincipal, adOpenDynamic, adLockBatchOptimistic
    LlenarLista
  End If
End Sub

Private Sub CmdNuevo_Click()
  II = 1
  FrmNuevoUsuario.Show 1
  LlenarLista
End Sub

Private Sub LlenarLista()
  LstUsuarios.ListItems.Clear
  AbrirRecorset rstUniversal, "Select*from Usuarios", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    Set Item = LstUsuarios.ListItems.Add(, , rstUniversal!IdUsuario)
      Item.SubItems(1) = rstUniversal!NmUsuario & ""
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
End Sub

Private Function DevSINO(Tip As Byte) As String
  If Tip = 0 Then
    DevSINO = "NO"
  Else
    DevSINO = "SI"
  End If
End Function
Private Function Dev10(SINO As String) As Byte
  If SINO = "SI" Then
    Dev10 = 1
  Else
    Dev10 = 0
  End If
End Function

Private Sub CmdPermisosEspeciales_Click()
  FufuLo = LstUsuarios.ListItems(LstUsuarios.SelectedItem.Index)
  FufuSt = LstUsuarios.ListItems(LstUsuarios.SelectedItem.Index).SubItems(1)
  FrmPermisosEspecialesUsuario.Show 1
End Sub

Private Sub CmdPermisosFormularios_Click()
  FufuLo = LstUsuarios.ListItems(LstUsuarios.SelectedItem.Index)
  FufuSt = LstUsuarios.ListItems(LstUsuarios.SelectedItem.Index).SubItems(1)
  FrmPermisos.Show 1
End Sub

Private Sub Form_Load()
  LlenarLista
End Sub
Private Sub LstUsuarios_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
    PopupMenu MnuPermisos
  End If
End Sub
Private Sub MnuVerPermisosRecogidas_Click()
  CargarPermisos 2
End Sub
Private Sub MnuVerPerMovimiento_Click()
  CargarPermisos 1
End Sub
Private Sub MnuVerPerVehiculos_Click()
  CargarPermisos 3
End Sub



