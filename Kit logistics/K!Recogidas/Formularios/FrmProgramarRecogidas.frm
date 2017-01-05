VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmProgramarRecogidas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programar recogidas..."
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   13050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdImprimirOrden 
      Caption         =   "Imprimir"
      Height          =   255
      Left            =   5880
      TabIndex        =   13
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton CmdCerrarRecogida 
      Caption         =   "Cerrar Recogida"
      Height          =   255
      Left            =   11040
      TabIndex        =   12
      Top             =   5040
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPFechaVehiculos 
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   7080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16908289
      CurrentDate     =   40634
   End
   Begin VB.CommandButton CmdAgregarAnunciosProgramados 
      Caption         =   "Anuncios programados"
      Height          =   255
      Left            =   11040
      TabIndex        =   10
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton CmdDescargasRemotas 
      Caption         =   "Descargas remotas"
      Height          =   255
      Left            =   7560
      TabIndex        =   8
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton CmdOpciones 
      Caption         =   "Opciones"
      Height          =   255
      Left            =   9480
      TabIndex        =   7
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton CmdAliminarRecogida 
      Caption         =   "Eliminar recogida"
      Height          =   255
      Left            =   11040
      TabIndex        =   4
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton CmdReProgramar 
      Caption         =   "Re-Programar"
      Height          =   255
      Left            =   11040
      TabIndex        =   3
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Timer TmActualizar 
      Interval        =   60000
      Left            =   480
      Top             =   0
   End
   Begin VB.CommandButton CmdAsignarMarcadas 
      Caption         =   "<< Asignar marcadas"
      Height          =   255
      Left            =   11040
      TabIndex        =   2
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton CmdAsignar 
      Caption         =   "<< Asignar"
      Height          =   255
      Left            =   11040
      TabIndex        =   1
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton CmdAct 
      Caption         =   "Ver pendientes"
      Height          =   255
      Left            =   11040
      TabIndex        =   0
      Top             =   3600
      Width           =   1815
   End
   Begin MSComctlLib.ListView LstVehiculos 
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "Imagenes"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Asignacion"
         Object.Width           =   1764
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
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Conductor"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ListView LstAnuncios 
      Height          =   3375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "Imagenes"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Anuncio"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cliente"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Anunciante"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ruta"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Hora"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Fecha Rec"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Direccion"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Unidades"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "KReal"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "KVol"
         Object.Width           =   1411
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
            Picture         =   "FrmProgramarRecogidas.frx":0000
            Key             =   "Ven"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgramarRecogidas.frx":2082
            Key             =   "Ok"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgramarRecogidas.frx":2A94
            Key             =   "Transito"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProgramarRecogidas.frx":2BEE
            Key             =   "Nor"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolVehiculos 
      Height          =   570
      Left            =   120
      TabIndex        =   9
      Top             =   7080
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   1005
      ButtonWidth     =   847
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuev"
            Object.ToolTipText     =   "Crear nuevo registro [F9]"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Object.ToolTipText     =   "Editar la informacion guardada [F10]"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Elim"
            Object.ToolTipText     =   "Elimina o anula el registro [F3]"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Act"
            Object.ToolTipText     =   "Actualizar la informacion"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ver"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Todas las recogidas del vehiculo"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Recogidas pendientes del vehiculo"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Recogidas "
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Auxiliares"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmProgramarRecogidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAct_Click()
  RecogidasPendientes
End Sub

Private Sub CmdAgregarAnunciosProgramados_Click()
  Dim rstRecProgramadas As ADODB.Recordset
  Set rstRecProgramadas = New ADODB.Recordset
  rstRecProgramadas.CursorLocation = adUseClient
  
  Dim rstTercero As ADODB.Recordset
  Set rstTercero = New ADODB.Recordset
  rstTercero.CursorLocation = adUseClient
  
  If MsgBox("¿Esta seguro de agregar los anuncios programados a los anuncios del dia?", vbQuestion + vbYesNo) = vbYes Then
    AbrirRecorset rstRecProgramadas, "Select anunciosprogramados.* from anunciosprogramados", CnnPrincipal, adOpenDynamic, adLockOptimistic
     Do While rstRecProgramadas.EOF = False
     AbrirRecorset rstTercero, "Select IdTercero, RazonSocial, Direccion, Telefono from terceros where IdTercero=" & rstRecProgramadas.Fields("IdCliente"), CnnPrincipal, adOpenDynamic, adLockOptimistic
     AbrirRecorset rstUniversal, "INSERT INTO Anuncios (IdAnuncio, FhAnuncio, IdCliente, Anunciante, DirAnunciante, TelAnunciante, IdRuta, FhRecogida, Unidades, KilosReales, KilosVol, Comentarios, Programada, Estado, Efectiva, Coperaciones, Orden, IdVehiculo, IdConductor, IdEmpresa) " & _
          " VALUES (" & SacarConsecutivo("Anuncios") & ", now() ,'" & rstRecProgramadas.Fields("IdCliente") & "','','" & rstTercero.Fields("Direccion") & "','" & rstTercero.Fields("Telefono") & "',Null,'" & Format(Date, "yyyy/mm/dd") & " " & Format(rstRecProgramadas.Fields("Hora"), "h:m:s") & "',1,1,1,'',0,'P',0," & Coperaciones & ",0,'','',1)", CnnPrincipal, adOpenDynamic, adLockOptimistic
      rstRecProgramadas.MoveNext
     Loop
    CerrarRecorset rstRecProgramadas
    RecogidasPendientes
  End If
End Sub

Private Sub CmdAliminarRecogida_Click()
On Error GoTo SinItem
  If MsgBox("¿Esta seguro de eliminar el anuncio de recogida?", vbYesNo + vbQuestion) = vbYes Then
    AbrirRecorset rstUniversal, "Delete from anuncios where IdAnuncio=" & LstAnuncios.SelectedItem, CnnPrincipal, adOpenDynamic, adLockBatchOptimistic
    LstAnuncios.ListItems.Remove LstAnuncios.SelectedItem.Index
    MsgBox "El anuncio de recogida ha sido eliminado con exito", vbInformation
  End If
SinItem:
  If Err.Number = 91 Then MsgBox "No hay un anuncio seleccionado", vbCritical
End Sub

Private Sub CmdAsignar_Click()
On Error GoTo SinItem
  If MsgBox("Se le va a agregar la recogida [" & LstAnuncios.ListItems(LstAnuncios.SelectedItem.Index) & "] al vehiculo [" & LstVehiculos.ListItems(LstVehiculos.SelectedItem.Index).SubItems(1) & "]", vbQuestion + vbYesNo) = vbYes Then
    AbrirRecorset rstUniversal, "Update anuncios set IdAsignacion=" & LstVehiculos.ListItems(LstVehiculos.SelectedItem.Index) & ", Programada=1, Estado='P' where IdAnuncio=" & LstAnuncios.ListItems(LstAnuncios.SelectedItem.Index), CnnPrincipal, adOpenDynamic, adLockOptimistic
    AbrirRecorset rstUniversal, "Update anuncios set orden=" & LstVehiculos.ListItems(LstVehiculos.SelectedItem.Index).SubItems(9) & " where IdAnuncio=" & LstAnuncios.ListItems(LstAnuncios.SelectedItem.Index), CnnPrincipal, adOpenDynamic, adLockOptimistic
    ResumirAsignacion Val(LstVehiculos.ListItems(LstVehiculos.SelectedItem.Index))
    LstAnuncios.ListItems.Remove LstAnuncios.SelectedItem.Index
    LlenarVehiculos
  End If
SinItem:
  If Err.Number = 91 Then MsgBox "No hay un vehiculo o un anuncio seleccionado", vbCritical
End Sub
Private Sub CmdAsignarMarcadas_Click()
If LstVehiculos.ListItems.Count > 0 Then
  II = 1
  Do While II <= LstAnuncios.ListItems.Count
    If LstAnuncios.ListItems(II).Checked = True Then
      AbrirRecorset rstUniversal, "Update anuncios set IdAsignacion=" & Val(LstVehiculos.ListItems(LstVehiculos.SelectedItem.Index)) & ", Programada=1, Estado='P' where IdAnuncio=" & Val(LstAnuncios.ListItems(II)), CnnPrincipal, adOpenDynamic, adLockOptimistic
      'FufuLo = Val(LstVehiculos.ListItems(LstVehiculos.SelectedItem.Index).SubItems(9)) + 1
      'AbrirRecorset rstUniversal, "Update anuncios set Orden=" & FufuLo & " where IdAnuncio=" & LstAnuncios.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
      LstAnuncios.ListItems.Remove II
    Else
      II = II + 1
    End If
  Loop
  ResumirAsignacion Val(LstVehiculos.ListItems(LstVehiculos.SelectedItem.Index))
  LlenarVehiculos
Else
  MsgBox "No hay vehiculos para asignarle recogidas", vbCritical
End If
End Sub

Private Sub CmdCerrarRecogida_Click()
On Error GoTo SinItem
  FufuLo = Val(LstAnuncios.SelectedItem)
  FrmCerrarAnuncio.Show 1
  RecogidasPendientes
SinItem:
  If Err.Number = 91 Then MsgBox "No hay un anuncio seleccionado", vbCritical
End Sub

Private Sub CmdDescargasRemotas_Click()
  FrmDescargasRemotas.Show 1
End Sub


Private Sub CmdImprimirOrden_Click()
On Error GoTo SinItem
  Mostrar_Reporte CnnPrincipal, 44, "select*from sql_ir_formato_recogidas where IdAsignacion=" & Val(LstVehiculos.SelectedItem), "Formato de recogidas", 2
SinItem:
  If Err.Number = 91 Then MsgBox "Debe seleccionar un vehiculo de ruta", vbCritical
End Sub

Private Sub CmdOpciones_Click()
  FrmOpciones.Show 1
End Sub


Private Sub CmdReProgramar_Click()
On Error GoTo SinItem
  Me.Tag = 1
  FufuSt = LstAnuncios.SelectedItem
  FrmReProgramarRec.DTPHora = LstAnuncios.ListItems(LstAnuncios.SelectedItem.Index).SubItems(4)
  FrmReProgramarRec.DTPFechaRe = LstAnuncios.ListItems(LstAnuncios.SelectedItem.Index).SubItems(5)
  FrmReProgramarRec.Show 1
  RecogidasPendientes
  Me.Tag = 0
SinItem:
  If Err.Number = 91 Then MsgBox "No hay un anuncio seleccionado", vbCritical
End Sub

Private Sub LlenarVehiculos()
  LstVehiculos.ListItems.Clear
  AbrirRecorset rstUniversal, "Select vehiculosrecogida.*, NmRuta, concat(Nombre, ' ', Apellido1, ' ', Apellido2) as NmConductor from vehiculosrecogida left join rutasurbanas on vehiculosrecogida.IdRuta=rutasurbanas.IdRutaRec left join conductores on vehiculosrecogida.IdConductor=conductores.IdConductor where Fecha='" & Format(DTPFechaVehiculos.Value, "yyyy-mm-dd") & "' and Coperaciones=" & Coperaciones, CnnPrincipal, adOpenDynamic, adLockOptimistic
  Do While rstUniversal.EOF = False
    If rstUniversal!Pend <= 0 Then
      Set Item = LstVehiculos.ListItems.Add(, , rstUniversal!IdAsignacion, , "Ok")
    Else
      Set Item = LstVehiculos.ListItems.Add(, , rstUniversal!IdAsignacion, , "Transito")
    End If
      Item.SubItems(1) = rstUniversal!Placa & ""
      Item.SubItems(2) = rstUniversal!Rec
      Item.SubItems(3) = rstUniversal!Pend
      Item.SubItems(4) = rstUniversal!Unidades
      Item.SubItems(5) = rstUniversal!KilosReales
      Item.SubItems(6) = rstUniversal!KilosVol
      Item.SubItems(7) = rstUniversal!NmRuta
      Item.SubItems(8) = rstUniversal!NmConductor & ""
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
End Sub

Private Sub RecogidasPendientes()
  II = 0
  LstAnuncios.ListItems.Clear
  AbrirRecorset rstUniversal, "SELECT Anuncios.IdAnuncio, Anuncios.Anunciante, Anuncios.DirAnunciante, Anuncios.IdRuta, Anuncios.FhRecogida, Anuncios.Unidades , Anuncios.KilosReales, Anuncios.KilosVol, Anuncios.Programada, Terceros.RazonSocial FROM anuncios left join terceros ON anuncios.IdCliente = terceros.IDTercero where Cerrada = 0 AND Programada=0 AND Coperaciones=" & Coperaciones, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    If rstUniversal!FhRecogida < Date Then II = 1
     If Format(rstUniversal!FhRecogida, "hh:mm") < Format(Time, "hh:mm") Then
       Set Item = LstAnuncios.ListItems.Add(, , rstUniversal!IdAnuncio, , "Ven")
     Else
       Set Item = LstAnuncios.ListItems.Add(, , rstUniversal!IdAnuncio, , "Nor")
     End If
     Item.SubItems(1) = rstUniversal!RazonSocial & ""
     Item.SubItems(2) = rstUniversal!Anunciante & ""
     Item.SubItems(3) = rstUniversal!IdRuta & ""
     Item.SubItems(4) = Format(rstUniversal!FhRecogida, "hh:mm")
     Item.SubItems(5) = Format(rstUniversal!FhRecogida, "dd-mm-yy")
     Item.SubItems(6) = rstUniversal!DirAnunciante & ""
     Item.SubItems(7) = rstUniversal!Unidades
     Item.SubItems(8) = rstUniversal!KilosReales
     Item.SubItems(9) = rstUniversal!KilosVol
     rstUniversal.MoveNext
  Loop
  If II = 1 Then
    MsgBox "Hay recogidas pendientes de fechas anteriores, debe eliminar o descargar estas recogidas", vbCritical
    'CmdAsignar.Enabled = False
    'CmdAsignarMarcadas.Enabled = False
    'ToolVehiculos.Buttons(3).Enabled = False
  Else
    CmdAsignar.Enabled = True
    CmdAsignarMarcadas.Enabled = True
    ToolVehiculos.Buttons(3).Enabled = True
  End If
  CerrarRecorset rstUniversal
  LstAnuncios.SetFocus
End Sub




Private Sub Form_Load()
  DTPFechaVehiculos.Value = Date
  LlenarVehiculos
  
  ToolVehiculos.ImageList = Principal.IgListTool
    ToolVehiculos.Buttons(3).Image = 1
    ToolVehiculos.Buttons(4).Image = 3
    ToolVehiculos.Buttons(5).Image = 4
    ToolVehiculos.Buttons(7).Image = 12
    ToolVehiculos.Buttons(8).Image = 14
  
End Sub
Private Sub LstAnuncios_KeyPress(KeyAscii As Integer)
  If CmdAsignar.Enabled = True Then If KeyAscii = 13 Then CmdAsignar_Click
End Sub
Private Sub LstAnuncios_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 8 Or KeyCode = 46 Then CmdAliminarRecogida_Click
End Sub

Private Sub TmActualizar_Timer()
  'If Val(Me.Tag) = 0 Then
  '  If MsgBox("¿Desea actualizar los anuncios pendientes?", vbQuestion + vbYesNo) = vbYes Then
  '    RecogidasPendientes
  '  End If
  'End If
End Sub

Private Sub ToolVehiculos_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index
    Case 3
      FufuLo = 0
      Me.Tag = 1
      FrmAgregarVehiculo.Show 1
      Me.Tag = 0
      If II = 1 Then
        LlenarVehiculos
      End If
    Case 4
      FufuLo = LstVehiculos.ListItems(LstVehiculos.SelectedItem.Index)
      FrmAgregarVehiculo.Show 1
      LlenarVehiculos
    Case 5
      On Error GoTo SinItem
      If Val(LstVehiculos.SelectedItem.SubItems(2)) = 0 Then
        AbrirRecorset rstUniversal, "Delete from VehiculosRecogida where IdAsignacion=" & Val(LstVehiculos.SelectedItem), CnnPrincipal, adOpenDynamic, adLockOptimistic
        AbrirRecorset rstUniversal, "Delete from AuxiliaresVehiculos where IdAsignacion=" & Val(LstVehiculos.SelectedItem), CnnPrincipal, adOpenDynamic, adLockOptimistic
        LstVehiculos.ListItems.Remove LstVehiculos.SelectedItem.Index
        MsgBox "Vehiculo eliminado con exito", vbInformation
      Else
        MsgBox "Este vehiculo tiene recogidas programadas, para sacarlo de la programacion primero debe quitarle las recogidas asignadas", vbCritical
      End If
    Case 7
      LlenarVehiculos
  End Select
  
SinItem:
  If Err.Number = 91 Then MsgBox "No hay un vehiculo seleccionado", vbCritical
    
End Sub

Private Sub ToolVehiculos_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo SinItem
  If ButtonMenu.Index = 4 Then
    FufuLo = LstVehiculos.SelectedItem
    FrmVerAuxiliares.Show 1
  Else
    Me.Tag = 1
    FufuSt = LstVehiculos.ListItems(LstVehiculos.SelectedItem.Index).SubItems(1)
    FufuLo = LstVehiculos.ListItems(LstVehiculos.SelectedItem.Index)
    II = ButtonMenu.Index - 1
    FrmVerRecogidas.Show 1
    ResumirAsignacion Val(LstVehiculos.ListItems(LstVehiculos.SelectedItem.Index))
    Me.Tag = 0
    LlenarVehiculos
  End If
SinItem:
  If Err.Number = 91 Then MsgBox "No hay un item seleccionado", vbCritical
End Sub
