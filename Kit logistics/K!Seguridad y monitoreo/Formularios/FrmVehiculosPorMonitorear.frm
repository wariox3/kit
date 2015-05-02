VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVehiculosPorMonitorear 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vehiculos pendientes por monitorear..."
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frecuencias de reporte"
      Height          =   615
      Left            =   7200
      TabIndex        =   5
      Top             =   5520
      Width           =   1935
      Begin VB.TextBox TxtFrec 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdCambiar 
         Caption         =   "Cambiar"
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdActualizar 
      Caption         =   "Actualizar"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton CmdCancelarMonitoreo 
      Caption         =   "Cancelar monitoreo"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton CmdIniciarMonitoreo 
      Caption         =   "Iniciar monitoreo"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   10440
      TabIndex        =   1
      Top             =   5760
      Width           =   1455
   End
   Begin MSComctlLib.ListView LstDespachos 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6376
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Orden"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tipo"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fecha"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Hora"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Vehiculo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Destino"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Frec"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView LstRecogidas 
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "Imagenes"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Asignacion"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Placa"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Rec"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Pend"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Unidades"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "KReal"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "KVol"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Ruta"
         Object.Width           =   4762
      EndProperty
   End
   Begin MSComctlLib.ImageList Imagenes 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVehiculosPorMonitorear.frx":0000
            Key             =   "Ven"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVehiculosPorMonitorear.frx":2082
            Key             =   "Ok"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVehiculosPorMonitorear.frx":2A94
            Key             =   "Transito"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVehiculosPorMonitorear.frx":2BEE
            Key             =   "Nor"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmVehiculosPorMonitorear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdActualizar_Click()
  VerDespachosSinMonitoreo
End Sub

Private Sub CmdCambiar_Click()
  If LstDespachos.ListItems.Count > 0 Then
    If Val(TxtFrec) > 0 Then
      LstDespachos.ListItems(LstDespachos.ListItems(LstDespachos.SelectedItem.Index).Index).SubItems(7) = Val(TxtFrec.Text)
      rstUniversal.Open "Update MonitoreoVehiculos set Frecuencia=" & Val(TxtFrec.Text) & " where ID=" & LstDespachos.SelectedItem, CnnPrincipal, adOpenDynamic, adLockOptimistic
      LstDespachos.SetFocus
    End If
  End If
End Sub
Private Sub CmdCancelarMonitoreo_Click()
  Dim rstActualizar As New ADODB.Recordset
  rstActualizar.CursorLocation = adUseClient
  II = 1
  While II <= LstDespachos.ListItems.Count
    If LstDespachos.ListItems(II).Checked = True Then
        rstActualizar.Open "UPDATE monitoreovehiculos SET SinMonitoreo=1 WHERE ID=" & Val(LstDespachos.ListItems.Item(II)), CnnPrincipal, adOpenDynamic, adLockOptimistic
        LstDespachos.ListItems.Remove (II)
    Else
     II = II + 1
    End If
  Wend
End Sub

Private Sub CmdIniciarMonitoreo_Click()
  Dim rstVehiculoRecogida As New ADODB.Recordset
  Dim FechaHora As String
  FechaHora = Format(Date, "yy/mm/dd") & " " & Format(Time, "h:m:s")
  rstVehiculoRecogida.CursorLocation = adUseClient
  II = 1
  While II <= LstDespachos.ListItems.Count
    If LstDespachos.ListItems(II).Checked = True Then
        AbrirRecorset rstUniversal, "UPDATE MonitoreoVehiculos SET Estado='T', UltReporte= '" & FechaHora & "' where ID=" & Val(LstDespachos.ListItems.Item(II)), CnnPrincipal, adOpenDynamic, adLockOptimistic
        LstDespachos.ListItems.Remove (II)
    Else
     II = II + 1
    End If
  Wend
  
  II = 1
  While II <= LstRecogidas.ListItems.Count
    If LstRecogidas.ListItems(II).Checked = True Then
      AbrirRecorset rstVehiculoRecogida, "SELECT IdAsignacion, Fecha from vehiculosrecogida WHERE IdAsignacion = " & LstRecogidas.SelectedItem, CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstVehiculoRecogida.RecordCount > 0 Then
        
        AbrirRecorset rstUniversal, "INSERT INTO monitoreovehiculos (Vehiculo, Destino, FhHrSalida, UltReporte, Estado, Tipo) VALUES('" & LstRecogidas.ListItems(II).SubItems(1) & "', 'RECOGIDA', '" & FechaHora & "', '" & FechaHora & "', 'T', 5)", CnnPrincipal, adOpenDynamic, adLockOptimistic
      End If
      CerrarRecorset rstVehiculoRecogida
    End If
    II = II + 1
  Wend
  VerDespachosSinMonitoreo
  LlenarVehiculos
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub
Private Sub Form_Load()
  VerDespachosSinMonitoreo
  LlenarVehiculos
End Sub

Private Sub VerDespachosSinMonitoreo()
  LstDespachos.ListItems.Clear
  AbrirRecorset rstUniversal, "Select*from MonitoreoVehiculos where Estado='P' and SinMonitoreo=0", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Do While rstUniversal.EOF = False
      Set Item = LstDespachos.ListItems.Add(, , rstUniversal.Fields("ID"))
        Item.SubItems(1) = rstUniversal.Fields("Orden")
        Item.SubItems(2) = DevTipo(rstUniversal.Fields("Tipo"))
        Item.SubItems(3) = Format(rstUniversal.Fields("FhHrSalida"), "dd/mm/yy")
        Item.SubItems(4) = Format(rstUniversal.Fields("FhHrSalida"), "HH:MM")
        Item.SubItems(5) = rstUniversal.Fields("Vehiculo")
        Item.SubItems(6) = rstUniversal.Fields("Destino")
        Item.SubItems(7) = rstUniversal.Fields("Frecuencia")
      rstUniversal.MoveNext
    Loop
  rstUniversal.Close
End Sub
Private Sub LstDespachos_ItemClick(ByVal Item As MSComctlLib.ListItem)
  TxtFrec.Text = LstDespachos.ListItems(LstDespachos.ListItems(LstDespachos.SelectedItem.Index).Index).SubItems(7)
End Sub
Private Sub LstDespachos_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then TxtFrec.SetFocus
End Sub

Private Sub TxtFrec_GotFocus()
  EnfocarT TxtFrec
End Sub

Private Sub TxtFrec_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then CmdCambiar.SetFocus
End Sub

Private Sub LlenarVehiculos()
  LstRecogidas.ListItems.Clear
  AbrirRecorset rstUniversal, "Select vehiculosrecogida.*, NmRuta from vehiculosrecogida left join rutasurbanas on vehiculosrecogida.IdRuta=rutasurbanas.IdRutaRec where Fecha='" & Format(Date, "yyyy-mm-dd") & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Do While rstUniversal.EOF = False
    If rstUniversal!Pend <= 0 Then
      Set Item = LstRecogidas.ListItems.Add(, , rstUniversal!IdAsignacion, , "Ok")
    Else
      Set Item = LstRecogidas.ListItems.Add(, , rstUniversal!IdAsignacion, , "Transito")
    End If
      Item.SubItems(1) = rstUniversal!Placa & ""
      Item.SubItems(2) = rstUniversal!Rec
      Item.SubItems(3) = rstUniversal!Pend
      Item.SubItems(4) = rstUniversal!Unidades
      Item.SubItems(5) = rstUniversal!KilosReales
      Item.SubItems(6) = rstUniversal!KilosVol
      Item.SubItems(7) = rstUniversal!NmRuta
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
End Sub
