VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBetrieb 
   Caption         =   "Vehiculos en monitoreo..."
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9915
   ScaleWidth      =   15705
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   255
      Left            =   8880
      TabIndex        =   21
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton CmdEliminarAcompañamiento 
      Caption         =   "Eliminar acompañamiento"
      Height          =   255
      Left            =   10440
      TabIndex        =   20
      Top             =   7320
      Width           =   2535
   End
   Begin VB.CommandButton CmdVerAcompañamientos 
      Caption         =   "Ver Acompañamientos"
      Height          =   255
      Left            =   10440
      TabIndex        =   19
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton CmdAgregarMonitoreo 
      Caption         =   "Agregar"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton CmdAbrirMonitoreo 
      Caption         =   "Abrir"
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   3960
      Width           =   1335
   End
   Begin MSComctlLib.ListView LstMonitoreos 
      Height          =   2055
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   3625
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IDC"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha y hora"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Control Post"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Notas"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Usuario"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.CommandButton CmdQuitarMonitoreo 
      Caption         =   "Quitar Reporte"
      Height          =   255
      Left            =   10440
      TabIndex        =   14
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton CmdAgregarAcompañamiento 
      Caption         =   "Agregar Acompañamiento"
      Height          =   255
      Left            =   10440
      TabIndex        =   13
      Top             =   6960
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   4320
      TabIndex        =   10
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton CmdCambiarFrecuencia 
         Caption         =   "Cambiar frecuencia"
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox TxtFrec 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdCerrarMonitoreo 
      Caption         =   "Cerrar"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton CmdNovedades 
      Caption         =   "Agregar/Solucionar novedad"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   3975
      Begin VB.Label Label3 
         Caption         =   "Alerta"
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   2280
         Picture         =   "FrmBetrieb.frx":0000
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Novedad"
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   1080
         Picture         =   "FrmBetrieb.frx":0442
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Normal"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   0
         Picture         =   "FrmBetrieb.frx":0884
         Top             =   120
         Width           =   480
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
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
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBetrieb.frx":0CC6
            Key             =   "Normal"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBetrieb.frx":1118
            Key             =   "Novedad"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBetrieb.frx":156A
            Key             =   "Alerta"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdAgregarReporte 
      Caption         =   "Agregar reporte"
      Height          =   255
      Left            =   10440
      TabIndex        =   3
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton CmdAnalizarTransito 
      Caption         =   "Refrescar"
      Height          =   255
      Left            =   10920
      TabIndex        =   2
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton CmdVerMonitoreoControlPost 
      Caption         =   "Refrescar"
      Height          =   255
      Left            =   10440
      TabIndex        =   1
      Top             =   5160
      Width           =   2415
   End
   Begin MSComctlLib.ListView LstDespachos 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
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
         Object.Width           =   2963
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Vehiculo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Destino"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Ult. Reporte"
         Object.Width           =   2963
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Frec"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Tiempo"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView LstAcompañamientos 
      Height          =   1455
      Left            =   120
      TabIndex        =   18
      Top             =   6600
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   2566
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
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IdAcompañante"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Acompañante"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Comentarios"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FrmBetrieb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAbrirMonitoreo_Click()
  Dim rstActualizar As New ADODB.Recordset
  rstActualizar.CursorLocation = adUseClient
  If Principal.ToolConsultas1.AbrirDevDatos("Numero de orden de despacho", "Digite el numero de la orden de despacho", 3, 0) = True Then
    AbrirRecorset rstUniversal, "Select*from MonitoreoVehiculos where Orden=" & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      rstActualizar.Open "Update MonitoreoVehiculos set Ok=0 where ID=" & rstUniversal.Fields("ID"), CnnPrincipal, adOpenDynamic, adLockOptimistic
      MsgBox "El monitoreo de este despacho se reanudo con exito", vbInformation
      CerrarRecorset rstUniversal
      AnalizarTransito
    Else
      MsgBox "No se encontraron monitoreos de esta orden de despacho", vbCritical
    End If
    CerrarRecorset rstUniversal
  End If
End Sub

Private Sub CmdAgregarAcompañamiento_Click()
  FufuLo = Val(LstDespachos.ListItems(LstDespachos.SelectedItem.Index))
  FrmAgregarEscolta.Show 1
  CmdVerAcompañamientos_Click
End Sub

Private Sub CmdAgregarMonitoreo_Click()
  FrmAgregarMonitoreo.Show 1
  AnalizarTransito
End Sub

Private Sub CmdAgregarReporte_Click()
  FufuLo = LstDespachos.ListItems(LstDespachos.SelectedItem.Index)
  FrmRegistroDeMonitoreo.Show 1
  CmdVerMonitoreoControlPost_Click
End Sub
Private Sub CmdAnalizarTransito_Click()
  AnalizarTransito
End Sub

Private Sub CmdCambiarFrecuencia_Click()
  If LstDespachos.ListItems.Count > 0 Then
    If Val(TxtFrec.Text) > 0 Then
      LstDespachos.ListItems(LstDespachos.ListItems(LstDespachos.SelectedItem.Index).Index).SubItems(7) = Val(TxtFrec.Text)
      rstUniversal.Open "Update MonitoreoVehiculos set Frecuencia=" & Val(TxtFrec.Text) & " where ID=" & LstDespachos.SelectedItem, CnnPrincipal, adOpenDynamic, adLockOptimistic
      LstDespachos.SetFocus
    Else
      MsgBox "La  frecuencia debe ser mayor a 0", vbCritical
    End If
  End If
End Sub

Private Sub CmdCerrarMonitoreo_Click()
  rstUniversal.Open "Update MonitoreoVehiculos set Ok=1 where ID=" & LstDespachos.ListItems(LstDespachos.SelectedItem.Index), CnnPrincipal, adOpenDynamic, adLockOptimistic
  MsgBox "El monitoreo de este despacho termino con exito", vbInformation
  AnalizarTransito
End Sub

Private Sub CmdEliminarAcompañamiento_Click()
  For II = 1 To LstAcompañamientos.ListItems.Count
    If LstAcompañamientos.ListItems(II).Checked = True Then
      rstUniversal.Open "Delete from monitoreo_acompañamiento where IdAcompañamiento=" & Val(LstAcompañamientos.ListItems(II).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
    End If
  Next
  CmdVerAcompañamientos_Click
End Sub

Private Sub CmdImprimir_Click()
  Mostrar_Reporte CnnPrincipal, 21, "Select*from sql_ism_monitoreos where IDMonitoreo=" & LstDespachos.ListItems(LstDespachos.SelectedItem.Index), "", 2
End Sub

Private Sub CmdNovedades_Click()
  FufuLo = LstDespachos.ListItems(LstDespachos.SelectedItem.Index)
  FrmNovedades.Show 1
  AnalizarTransito
End Sub

Private Sub CmdQuitarMonitoreo_Click()
  If CpPermisoEspecial(18, CodUsuarioActivo, CnnPrincipal) = True Then
    For II = 1 To LstMonitoreos.ListItems.Count
      If LstMonitoreos.ListItems(II).Checked = True Then
        rstUniversal.Open "Delete from MonitoreoControlPost where Id=" & LstMonitoreos.ListItems(II).Text, CnnPrincipal, adOpenDynamic, adLockOptimistic
      End If
    Next
    CmdVerMonitoreoControlPost_Click
  Else
    MsgBox "No tiene permisos para esta opcion", vbInformation
  End If
End Sub

Private Sub CmdVerAcompañamientos_Click()
  LstAcompañamientos.ListItems.Clear
  rstUniversal.Open "SELECT monitoreo_acompañamiento.*, terceros.RazonSocial FROM monitoreo_acompañamiento left join terceros ON monitoreo_acompañamiento.IdEscolta = terceros.IdTercero Where IdMonitoreo=" & Val(LstDespachos.ListItems(LstDespachos.SelectedItem.Index)), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Do While rstUniversal.EOF = False
      Set Item = LstAcompañamientos.ListItems.Add(, , rstUniversal.Fields("IdAcompañamiento"))
      Item.SubItems(1) = rstUniversal.Fields("IdEscolta")
      Item.SubItems(2) = rstUniversal.Fields("RazonSocial")
      Item.SubItems(3) = rstUniversal.Fields("VrAcompañamiento") & ""
      Item.SubItems(4) = rstUniversal.Fields("ComentariosAcompañamiento") & ""
      rstUniversal.MoveNext
    Loop
  rstUniversal.Close
End Sub

Private Sub CmdVerMonitoreoControlPost_Click()
  LstMonitoreos.ListItems.Clear
  If (LstDespachos.ListItems.Count > 0) Then
    rstUniversal.Open "SELECT monitoreocontrolpost.*, ControlPost.NmControlPost FROM monitoreocontrolpost INNER JOIN ControlPost ON monitoreocontrolpost.IDControlPost = ControlPost.IdControlPost Where IdMonitoreo=" & LstDespachos.ListItems(LstDespachos.SelectedItem.Index), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      Do While rstUniversal.EOF = False
        Set Item = LstMonitoreos.ListItems.Add(, , rstUniversal.Fields("ID"))
        Item.SubItems(1) = rstUniversal.Fields("IdControlPost")
        Item.SubItems(2) = Format(rstUniversal.Fields("FhHrReporte"), "dd/mm/yy HH:mm AM/PM")
        Item.SubItems(3) = rstUniversal.Fields("NmControlPost") & ""
        Item.SubItems(4) = rstUniversal.Fields("Notas") & ""
        Item.SubItems(5) = rstUniversal.Fields("usuario") & ""
        rstUniversal.MoveNext
      Loop
    rstUniversal.Close
  End If
  
End Sub
Private Sub Form_Load()
  AnalizarTransito
End Sub
Private Sub AnalizarTransito()
Dim FH1 As Date, FH2 As Date
  LstDespachos.ListItems.Clear
  AbrirRecorset rstUniversal, "select*from MonitoreoVehiculos where Estado='T' and Ok=0", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Do While rstUniversal.EOF = False
      If rstUniversal.Fields("UltReporte") <> "" Then
        FH1 = DateAdd("n", rstUniversal.Fields("Frecuencia"), rstUniversal.Fields("UltReporte"))
      End If
      FH2 = Date & " " & Time
      If Val(rstUniversal.Fields("EnNovedad")) = 0 Then
        If FH1 > FH2 Then
          Set Item = LstDespachos.ListItems.Add(, , rstUniversal.Fields("ID"), "Normal", "Normal")
        Else
          Set Item = LstDespachos.ListItems.Add(, , rstUniversal.Fields("ID"), "Alerta", "Alerta")
        End If
      Else
        Set Item = LstDespachos.ListItems.Add(, , rstUniversal.Fields("ID"), , "Novedad")
      End If
      Item.SubItems(1) = rstUniversal.Fields("Orden")
      Item.SubItems(2) = DevTipo(Val(rstUniversal.Fields("Tipo") & ""))
      Item.SubItems(3) = Format(rstUniversal.Fields("FhHrSalida"), "dd/mm/yy HH:mm AM/PM")
      Item.SubItems(4) = rstUniversal.Fields("Vehiculo")
      Item.SubItems(5) = rstUniversal.Fields("Destino")
      Item.SubItems(6) = Format(rstUniversal.Fields("UltReporte"), "dd/mm/yy HH:mm AM/PM")
      Item.SubItems(7) = rstUniversal.Fields("Frecuencia")
      rstUniversal.MoveNext
    Loop
  rstUniversal.Close
End Sub

Private Sub LstDespachos_ItemClick(ByVal Item As MSComctlLib.ListItem)
  CmdVerMonitoreoControlPost_Click
  CmdVerAcompañamientos_Click
End Sub

Private Sub TxtFrec_GotFocus()
  EnfocarT TxtFrec
End Sub

Private Sub TxtFrec_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then CmdCambiarFrecuencia.SetFocus
End Sub
