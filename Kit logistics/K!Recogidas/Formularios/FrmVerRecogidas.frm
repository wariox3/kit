VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVerRecogidas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ver recogidas..."
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBajar 
      Height          =   615
      Left            =   11400
      Picture         =   "FrmVerRecogidas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton CmdSubir 
      Height          =   615
      Left            =   11400
      Picture         =   "FrmVerRecogidas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton CmdVerReporte 
      Caption         =   "&Ver reporte"
      Height          =   255
      Left            =   7800
      TabIndex        =   10
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton CmdCargarRecogida 
      Caption         =   "Cargar recogida"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton CmdDescargar 
      Caption         =   "Descargar"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton CmdQuitarMarcadas 
      Caption         =   "<< Quitar Marcadas"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton CmdQuitar 
      Caption         =   "<< Quitar"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   9360
      TabIndex        =   1
      Top             =   4680
      Width           =   1935
   End
   Begin MSComctlLib.ListView LstAnuncios 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7435
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Anuncio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Anunciante / Cliente"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Ruta"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Hora"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Fecha Rec"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Direccion"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Unidades"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "KReal"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "KVol"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Ord"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label LblAsiganacion 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   4200
      TabIndex        =   8
      Top             =   0
      Width           =   1050
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Asignacion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   2760
      TabIndex        =   7
      Top             =   0
      Width           =   1410
   End
   Begin VB.Label LblVehiculo 
      AutoSize        =   -1  'True
      Caption         =   "TIX125"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   0
      Width           =   885
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Vehiculo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   1125
   End
End
Attribute VB_Name = "FrmVerRecogidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ItemTem As ListItem
Private Sub CmdCargarRecogida_Click()
On Error GoTo SinItem
  AbrirRecorset rstUniversal, "Update Anuncios set Efectiva=0 where IdAnuncio=" & LstAnuncios.SelectedItem, CnnPrincipal, adOpenDynamic, adLockOptimistic
  LstAnuncios.ListItems.Remove LstAnuncios.SelectedItem.Index
  FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(3) = Val(FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(3)) + 1
  MsgBox "El anuncio se volvio a cargar con exito, ahora quedo pendiente", vbInformation
SinItem:
  If Err.Number = 91 Then MsgBox "No hay un anuncio seleccionado", vbCritical
End Sub

Private Sub CmdDescargar_Click()
On Error GoTo SinItem
  FrmDevDescargar.TxtUnidades = Val(LstAnuncios.ListItems(LstAnuncios.SelectedItem.Index).SubItems(6))
  FrmDevDescargar.TxtKReales = Val(LstAnuncios.ListItems(LstAnuncios.SelectedItem.Index).SubItems(7))
  FrmDevDescargar.TxtKVol = Val(LstAnuncios.ListItems(LstAnuncios.SelectedItem.Index).SubItems(8))
  FrmDevDescargar.DTPHora = LstAnuncios.ListItems(LstAnuncios.SelectedItem.Index).SubItems(3)
  FrmDevDescargar.Show 1
SinItem:
  If Err.Number = 91 Then MsgBox "No hay un anuncio seleccionado", vbCritical
End Sub

Private Sub Encabezado()
  UbC 1, 3, 5, 12, "ASIGNACION [" & LblAsiganacion.Caption & "]  VEHICULO [" & LblVehiculo.Caption & "]   FECHA [" & Format(Date, "dd/mmm/yyyy") & "]", 100
  UbC 1, 3, 15, 10, "ANUNCIO CLIENTE                               HORA     UND           KR            KV", 100

End Sub
Private Sub CmdQuitar_Click()
On Error GoTo SinItem
  If MsgBox("Le va a quitar la recogida [" & LstAnuncios.ListItems(LstAnuncios.SelectedItem.Index).SubItems(1) & "] al vehiculo [" & LblVehiculo.Caption & "]", vbQuestion + vbYesNo) = vbYes Then
    AbrirRecorset rstUniversal, "Update Anuncios set IdAsignacion=0, Programada=0, Estado='I' where IdAnuncio=" & LstAnuncios.ListItems(LstAnuncios.SelectedItem.Index), CnnPrincipal, adOpenDynamic, adLockOptimistic
    FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(2) = Val(FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(2)) - 1
    FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(3) = Val(FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(3)) - 1
    FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(4) = Val(FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(4)) - Val(LstAnuncios.ListItems(LstAnuncios.SelectedItem.Index).SubItems(6))
    FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(5) = Val(FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(5)) - Val(LstAnuncios.ListItems(LstAnuncios.SelectedItem.Index).SubItems(7))
    FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(6) = Val(FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(6)) - Val(LstAnuncios.ListItems(LstAnuncios.SelectedItem.Index).SubItems(8))
    LstAnuncios.ListItems.Remove LstAnuncios.SelectedItem.Index
  Else
    LstAnuncios.ListItems(II).Checked = False
  End If
SinItem:
  If Err.Number = 91 Then MsgBox "No hay un anuncio seleccionado", vbCritical
End Sub

Private Sub CmdQuitarMarcadas_Click()
II = 1
Do While II <= LstAnuncios.ListItems.Count
  If LstAnuncios.ListItems(II).Checked = True Then
    AbrirRecorset rstUniversal, "Update Anuncios set IdAsignacion=0, Programada=0, Estado='I' where IdAnuncio=" & LstAnuncios.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
        
    LstAnuncios.ListItems.Remove II
  Else
    II = II + 1
  End If
Loop
ResumirAsignacion (Val(LblAsiganacion.Caption))
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub CmdSubir_Click()
  If LstAnuncios.SelectedItem.Index > 1 Then
    FufuLo = LstAnuncios.SelectedItem.Index
    
    LstAnuncios.ListItems(FufuLo).Text = LstAnuncios.ListItems(FufuLo - 1).Text
    For II = 1 To 9
      LstAnuncios.ListItems(FufuLo).SubItems(II) = LstAnuncios.ListItems(FufuLo - 1).SubItems(II)
    Next
    MsgBox ItemTem.ListSubItems(1)
  End If
End Sub

Private Sub CmdVerReporte_Click()
  Select Case Val(Me.Tag)
    Case 0
      Mostrar_Reporte CnnPrincipal, 23, "select*from sql_ir_listadorecogidasvehiculo where IdAsignacion=" & Val(LblAsiganacion), "RECOGIDAS DE LA ASIGNACION", 2
    Case 1
      Mostrar_Reporte CnnPrincipal, 23, "select*from sql_ir_listadorecogidasvehiculo where Efectiva=0 and IdAsignacion=" & Val(LblAsiganacion), "RECOGIDAS PENDIENTES DE LA ASIGNACION", 2
    Case 2
      Mostrar_Reporte CnnPrincipal, 23, "select*from sql_ir_listadorecogidasvehiculo where Efectiva=1 and IdAsignacion=" & Val(LblAsiganacion), "RECOGIDAS EFECTIVAS/DESCARGADAS DE LA ASIGNACION", 2
  End Select
End Sub

Private Sub Form_Load()
  Me.Tag = II
  LblVehiculo.Caption = FufuSt
  LblAsiganacion.Caption = FufuLo
  Select Case Val(Me.Tag)
    Case 0
      AbrirRecorset rstUniversal, "SELECT IdAnuncio, Anunciante, DirAnunciante, IdRuta, FhRecogida, Unidades, KilosReales, KilosVol, Programada, Orden, RazonSocial From anuncios left join terceros on anuncios.IdCliente=terceros.IDTercero where IdAsignacion=" & Val(LblAsiganacion.Caption) & " Order by Orden", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      CmdBajar.Visible = True
      CmdSubir.Visible = True
    Case 1
      AbrirRecorset rstUniversal, "SELECT IdAnuncio, Anunciante, DirAnunciante, IdRuta, FhRecogida, Unidades, KilosReales, KilosVol, Programada, Orden, RazonSocial From anuncios left join terceros on anuncios.IdCliente=terceros.IDTercero where IdAsignacion=" & Val(LblAsiganacion.Caption) & " and Efectiva=0 Order by Orden", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      CmdQuitar.Visible = True
      CmdQuitarMarcadas.Visible = True
      CmdDescargar.Visible = True
    Case 2
      AbrirRecorset rstUniversal, "SELECT IdAnuncio, Anunciante, DirAnunciante, IdRuta, FhRecogida, Unidades, KilosReales, KilosVol, Programada,Orden, RazonSocial From anuncios left join terceros on anuncios.IdCliente=terceros.IDTercero where IdAsignacion=" & Val(LblAsiganacion.Caption) & " and Efectiva=1 Order by Orden", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      CmdCargarRecogida.Visible = True
  End Select
  Do While rstUniversal.EOF = False
    Set Item = LstAnuncios.ListItems.Add(, , rstUniversal!IdAnuncio)
      Item.SubItems(1) = rstUniversal!RazonSocial & ""
      Item.SubItems(2) = rstUniversal!IdRuta & ""
      Item.SubItems(3) = Format(rstUniversal!FhRecogida, "hh:mm")
      Item.SubItems(4) = Format(rstUniversal!FhRecogida, "dd/mm/yy")
      Item.SubItems(5) = rstUniversal!DirAnunciante & ""
      Item.SubItems(6) = rstUniversal!Unidades
      Item.SubItems(7) = rstUniversal!KilosReales
      Item.SubItems(8) = rstUniversal!KilosVol
      Item.SubItems(9) = rstUniversal!Orden
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
End Sub


Private Sub LstAnuncios_KeyPress(KeyAscii As Integer)
  If Val(Me.Tag) = 1 Then
    If KeyAscii = 13 Then CmdDescargar_Click
  End If
End Sub
