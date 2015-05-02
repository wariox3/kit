VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVerMonitoreos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ver Monitoreos"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LstNovedades 
      Height          =   1455
      Left            =   120
      TabIndex        =   21
      Top             =   4440
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImListLista"
      SmallIcons      =   "ImListLista"
      ColHdrIcons     =   "ImListLista"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "UsuIng"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fh/Hr Ingreso"
         Object.Width           =   2910
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fh/Hr Novedad"
         Object.Width           =   2910
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Novedad"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Comentarios"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "UsuSol"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Fh/Hr Sol"
         Object.Width           =   2910
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Solucion"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton CmdVerInfoDespacho 
      Caption         =   "Ver informacion del despacho"
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   7560
      Width           =   2535
   End
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10095
      Begin VB.TextBox TxtDestino 
         Height          =   285
         Left            =   7320
         TabIndex        =   19
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox TxtFhHrSalida 
         Height          =   285
         Left            =   7320
         TabIndex        =   18
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox TxtUltReporte 
         Height          =   285
         Left            =   7320
         TabIndex        =   17
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox TxtVehiculo 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox TxtEstado 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TxtDespacho 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxtId 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox ChkEnNovedad 
         Caption         =   "En novedad"
         Height          =   255
         Left            =   7320
         TabIndex        =   12
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox ChkTerminado 
         Caption         =   "Terminado"
         Height          =   255
         Left            =   8640
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Ult Reporte:"
         Height          =   195
         Index           =   6
         Left            =   6360
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Vehiculo:"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   10
         Top             =   1320
         Width           =   660
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   4
         Left            =   6630
         TabIndex        =   9
         Top             =   960
         Width           =   585
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Fh Hr Salida:"
         Height          =   195
         Index           =   3
         Left            =   6300
         TabIndex        =   8
         Top             =   600
         Width           =   915
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Despacho:"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   780
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   5
         Top             =   960
         Width           =   540
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         Height          =   195
         Index           =   0
         Left            =   930
         TabIndex        =   4
         Top             =   240
         Width           =   210
      End
   End
   Begin VB.CommandButton CmdActualizar 
      Caption         =   "Actualizar"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   255
      Left            =   8520
      TabIndex        =   1
      Top             =   7560
      Width           =   1695
   End
   Begin MSComctlLib.ListView LstMonitoreos 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   4260
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
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Notas"
         Object.Width           =   6174
      EndProperty
   End
   Begin MSComctlLib.ImageList ImListLista 
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
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerMonitoreos.frx":0000
            Key             =   "Ok"
            Object.Tag             =   "Ok"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerMonitoreos.frx":0A14
            Key             =   "Pendiente"
            Object.Tag             =   "Pendiente"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LstAcompañamientos 
      Height          =   1455
      Left            =   120
      TabIndex        =   22
      Top             =   6000
      Width           =   10095
      _ExtentX        =   17806
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
Attribute VB_Name = "FrmVerMonitoreos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstMonitoreos As New ADODB.Recordset

Private Sub CmdActualizar_Click()
  FufuLo = Val(TxtDespacho.Text)
  LlenarLista
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub CmdVerInfoDespacho_Click()
  FufuLo = Val(TxtDespacho.Text)
  FrmInfoDespacho.Show 1
End Sub

Private Sub Form_Load()
  rstMonitoreos.CursorLocation = adUseClient
  LlenarLista
  VerNovedades
  VerAcompañamientos
End Sub
Sub VerNovedades()
  LstNovedades.ListItems.Clear
  rstUniversal.Open "SELECT NovedadesMonitoreo.*, CausalesNovedadMonitoreo.NmNovedad FROM NovedadesMonitoreo INNER JOIN CausalesNovedadMonitoreo ON NovedadesMonitoreo.IdNovedad = CausalesNovedadMonitoreo.IdNovedad Where IdMonitoreo=" & Val(TxtID.Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    If Val(rstUniversal.Fields("Solucionada")) = 0 Then
      Set Item = LstNovedades.ListItems.Add(, , rstUniversal!Id, "Pendiente", "Pendiente")
    Else
      Set Item = LstNovedades.ListItems.Add(, , rstUniversal!Id, "Ok", "Ok")
    End If
    Item.SubItems(1) = rstUniversal!UsuIng & ""
    Item.SubItems(2) = Format(rstUniversal!FHIngreso, "dd/mmm/yy hh:mm:ss")
    Item.SubItems(3) = Format(rstUniversal!FHNovedad, "dd/mmm/yy hh:mm:ss")
    Item.SubItems(4) = rstUniversal!NmNovedad & ""
    Item.SubItems(5) = rstUniversal!Comentarios & ""
    Item.SubItems(6) = rstUniversal!UsuSol & ""
    Item.SubItems(7) = Format(rstUniversal!FHSolucion, "dd/mmm/yy hh:mm:ss")
    Item.SubItems(8) = rstUniversal!Solucion & ""
    rstUniversal.MoveNext
  Loop
  rstUniversal.Close
End Sub

Private Sub LlenarLista()
  rstMonitoreos.Open "Select*from monitoreovehiculos where Orden=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstMonitoreos.EOF = False Then
    TxtID.Text = rstMonitoreos.Fields("Id")
    TxtDespacho.Text = rstMonitoreos.Fields("Orden")
    TxtEstado.Text = rstMonitoreos.Fields("Estado")
    TxtFhHrSalida.Text = rstMonitoreos.Fields("FhHrSalida") & ""
    TxtVehiculo.Text = rstMonitoreos.Fields("Vehiculo") & ""
    TxtDestino.Text = rstMonitoreos.Fields("Destino") & ""
    TxtUltReporte.Text = rstMonitoreos.Fields("UltReporte") & ""
    TxtVehiculo.Text = rstMonitoreos.Fields("Vehiculo") & ""
    ChkEnNovedad.value = DevCheck(rstMonitoreos.Fields("EnNovedad"))
    ChkTerminado.value = DevCheck(rstMonitoreos.Fields("Ok"))
    
    rstMonitoreos.Close
    LstMonitoreos.ListItems.Clear
    rstMonitoreos.Open "SELECT MonitoreoControlPost.*, ControlPost.NmControlPost FROM MonitoreoControlPost INNER JOIN ControlPost ON MonitoreoControlPost.IDControlPost = ControlPost.IdControlPost Where IdMonitoreo=" & Val(TxtID.Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      Do While rstMonitoreos.EOF = False
        Set Item = LstMonitoreos.ListItems.Add(, , rstMonitoreos.Fields("ID"))
        Item.SubItems(1) = rstMonitoreos.Fields("IdControlPost")
        Item.SubItems(2) = Format(rstMonitoreos.Fields("FhHrReporte"), "dd/mm/yy HH:mm AM/PM")
        Item.SubItems(3) = rstMonitoreos.Fields("NmControlPost") & ""
        Item.SubItems(4) = rstMonitoreos.Fields("Notas") & ""
        rstMonitoreos.MoveNext
      Loop
    rstMonitoreos.Close
  Else
    rstMonitoreos.Close
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstMonitoreos = Nothing
End Sub

Private Sub VerAcompañamientos()
  Dim rstAcompañamientos As New ADODB.Recordset
  rstAcompañamientos.CursorLocation = adUseClient
  LstAcompañamientos.ListItems.Clear
  rstAcompañamientos.Open "SELECT monitoreo_acompañamiento.*, terceros.RazonSocial FROM monitoreo_acompañamiento left join terceros ON monitoreo_acompañamiento.IdEscolta = terceros.IdTercero Where IdMonitoreo=" & Val(TxtID.Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Do While rstAcompañamientos.EOF = False
      Set Item = LstAcompañamientos.ListItems.Add(, , rstAcompañamientos.Fields("IdAcompañamiento"))
      Item.SubItems(1) = rstAcompañamientos.Fields("IdEscolta")
      Item.SubItems(2) = rstAcompañamientos.Fields("RazonSocial")
      Item.SubItems(3) = rstAcompañamientos.Fields("VrAcompañamiento") & ""
      Item.SubItems(4) = rstAcompañamientos.Fields("ComentariosAcompañamiento") & ""
      rstAcompañamientos.MoveNext
    Loop
  rstAcompañamientos.Close
End Sub
