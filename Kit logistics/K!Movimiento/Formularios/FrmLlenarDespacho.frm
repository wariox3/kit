VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmLlenarDespacho 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "llenar despacho..."
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   13335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAgregarPorDocumento 
      Caption         =   "&Agregar por documento"
      Height          =   255
      Left            =   2160
      TabIndex        =   37
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton CmdAgregarBandejaAnulacion 
      Caption         =   "Agregar guias del ultimo manifiesto anulado"
      Height          =   255
      Left            =   6960
      TabIndex        =   36
      Top             =   3600
      Width           =   3495
   End
   Begin VB.CommandButton CmdPendientesPorImprimir 
      Caption         =   "Guias sin imprimir"
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   7440
      Width           =   2655
   End
   Begin VB.Frame FraPendientes 
      Height          =   3375
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   13095
      Begin VB.TextBox TxtIdCiudad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   34
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox TxtIdRuta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   33
         Top             =   2520
         Width           =   495
      End
      Begin VB.CommandButton CmdAgregarALista 
         Caption         =   "Agregar a lista"
         Height          =   255
         Left            =   9000
         TabIndex        =   28
         Top             =   2520
         Width           =   1935
      End
      Begin VB.OptionButton OptTodos 
         Alignment       =   1  'Right Justify
         Caption         =   "Todos"
         Height          =   255
         Left            =   5280
         TabIndex        =   26
         Top             =   2760
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton CmdCargarPendientes 
         Caption         =   "Cargar Pendientes"
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
         Left            =   11040
         TabIndex        =   23
         Top             =   2520
         Width           =   1935
      End
      Begin VB.OptionButton OptRuta 
         Alignment       =   1  'Right Justify
         Caption         =   "Ruta"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2520
         Width           =   735
      End
      Begin VB.OptionButton OptCiudad 
         Alignment       =   1  'Right Justify
         Caption         =   "Ciudad"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   2910
         Width           =   855
      End
      Begin VB.TextBox TxtNroResultados 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   12360
         TabIndex        =   18
         Text            =   "100"
         Top             =   2880
         Width           =   615
      End
      Begin MSDataGridLib.DataGrid GrillaResultados 
         Height          =   2175
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   3836
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "Guia"
            Caption         =   "Guia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   9226
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "COIng"
            Caption         =   "CI"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   9226
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Cliente"
            Caption         =   "Cliente"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   9226
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "DocCliente"
            Caption         =   "Documento"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   9226
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "NmDestinatario"
            Caption         =   "Destinatario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   9226
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "NmCiudad"
            Caption         =   "Destino"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   9226
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Unidades"
            Caption         =   "Unidades"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   9226
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "KilosReales"
            Caption         =   "K Real"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   9226
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "KilosVolumen"
            Caption         =   "K Vol"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   9226
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "IdRuta"
            Caption         =   "Ruta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   9226
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   404.787
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1635.024
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   2385.071
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   2520
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               ColumnWidth     =   540.284
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo CboCiudad 
         Height          =   315
         Left            =   1080
         TabIndex        =   22
         Top             =   2880
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo CboRuta 
         Height          =   315
         Left            =   1080
         TabIndex        =   24
         Top             =   2520
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nro Resultados:"
         Height          =   195
         Left            =   11160
         TabIndex        =   25
         Top             =   2880
         Width           =   1140
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar / Salir"
      Height          =   375
      Left            =   10560
      TabIndex        =   15
      Top             =   7440
      Width           =   2655
   End
   Begin VB.CommandButton CmdQuitar 
      Caption         =   "&Retirar guia"
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Totales"
      Enabled         =   0   'False
      Height          =   3375
      Left            =   10560
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   29
         Top             =   1266
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   2292
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   3
         Top             =   1950
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   4
         Top             =   1608
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   5
         Top             =   582
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   6
         Top             =   924
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   6
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   8
         Left            =   1200
         TabIndex        =   31
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "T Recaudo:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "K Vol:"
         Height          =   195
         Index           =   0
         Left            =   765
         TabIndex        =   30
         Top             =   1266
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "K Reales:"
         Height          =   195
         Index           =   2
         Left            =   495
         TabIndex        =   13
         Top             =   924
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unidades:"
         Height          =   195
         Index           =   3
         Left            =   465
         TabIndex        =   12
         Top             =   582
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remesas:"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cobro destino:"
         Height          =   195
         Index           =   6
         Left            =   150
         TabIndex        =   10
         Top             =   2295
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Flete:"
         Height          =   195
         Index           =   4
         Left            =   795
         TabIndex        =   9
         Top             =   1608
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Manejo:"
         Height          =   195
         Index           =   8
         Left            =   615
         TabIndex        =   8
         Top             =   1950
         Width           =   570
      End
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar por guia"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   1935
   End
   Begin MSComctlLib.ListView LstRem 
      Height          =   3375
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Guia"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "FhEntrada"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cliente"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Destino"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Unidades"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "KReales"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "K Vol"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Flete"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Manejo"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Recaudo"
         Object.Width           =   2029
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "TipoCobro"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "FrmLlenarDespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstTem As New ADODB.Recordset

Private Sub CboCiudad_GotFocus()
  AbrirRecorset rstUniversal, "Select IdCiudad, NmCiudad from Ciudades order by NmCiudad", CnnPrincipal, adOpenDynamic, adLockOptimistic
  CboCiudad.ListField = "NmCiudad"
  Set CboCiudad.RowSource = rstUniversal
End Sub

Private Sub CboCiudad_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then CmdCargarPendientes.SetFocus
End Sub

Private Sub CboCiudad_Validate(Cancel As Boolean)
  AbrirRecorset rstUniversalSer, "Select IdCiudad, NmCiudad from Ciudades where NmCiudad='" & CboCiudad & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversalSer.EOF = False Then
    CboCiudad.Tag = rstUniversalSer!IdCiudad
    TxtIdCiudad.Text = rstUniversalSer!IdCiudad
  Else
    CboCiudad.Tag = "": TxtIdCiudad.Text = ""
  End If
  CerrarRecorset rstUniversalSer
End Sub

Private Sub CboRuta_GotFocus()
  AbrirRecorset rstUniversal, "Select IdRuta, NmRuta from Rutas order by NmRuta", CnnPrincipal, adOpenDynamic, adLockOptimistic
  CboRuta.ListField = "NmRuta"
  Set CboRuta.RowSource = rstUniversal
End Sub

Private Sub CboRuta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then CmdCargarPendientes.SetFocus
End Sub

Private Sub CboRuta_Validate(Cancel As Boolean)
  AbrirRecorset rstUniversalSer, "Select IdRuta, NmRuta from Rutas where NmRuta='" & CboRuta & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversalSer.EOF = False Then
    CboRuta.Tag = rstUniversalSer!IdRuta
    TxtIdRuta.Text = rstUniversalSer!IdRuta
  Else
    CboRuta.Tag = "": TxtIdRuta.Text = ""
  End If
  CerrarRecorset rstUniversalSer
End Sub
Private Sub CmdAceptar_Click()
  AbrirRecorset rstUniversal, "Update Despachos set Unidades=" & Val(TxtTotales(4)) & ", KilosReales=" & Val(TxtTotales(5)) & " where OrdDespacho=" & Val(Me.Tag), CnnPrincipal, adOpenDynamic, adLockOptimistic
  FrmManifiestos.TxtCampos(11) = Val(Format(TxtTotales(6).Text, "0;(0)"))
  FrmManifiestos.TxtCampos(12) = Val(Format(TxtTotales(4).Text, "0;(0)"))
  FrmManifiestos.TxtCampos(13) = Val(Format(TxtTotales(5).Text, "0;(0)"))
  FrmManifiestos.TxtCampos(14) = Val(Format(TxtTotales(7).Text, "0;(0)"))
  FrmManifiestos.TxtCampos(15) = Val(Format(TxtTotales(3).Text, "0;(0)"))
  FrmManifiestos.TxtCampos(16) = Val(Format(TxtTotales(2).Text, "0;(0)"))
  FrmManifiestos.TxtTotalCE = Val(Format(TxtTotales(1).Text, "0;(0)"))
  'FrmManifiestos.TxtCampos(18) = Val(Format(TxtTotales(0).Text, "0;(0)"))
  FrmManifiestos.TxtCampos(33) = Val(Format(TxtTotales(8).Text, "0;(0)"))
  FrmManifiestos.TxtCampos(34) = Val(Format(TxtTotales(9).Text, "0;(0)"))
  CerrarRecorset rstUniversal
  Unload Me
End Sub

Private Sub CmdAgregar_Click()
  Do While Principal.ToolConsultas1.AbrirDevDatos("Digite el numero de Guia", "Digite el numero de guia para agregarle al despacho", 3, 0) = True
    FufuLo = Principal.ToolConsultas1.DatLo
    AgregarGuia FufuLo
  Loop
End Sub

Private Sub CmdAgregarALista_Click()
  If rstTem.State = adStateOpen Then
    If rstTem.RecordCount > 0 Then
      AgregarGuia rstTem.Fields("Guia")
      USiguiente rstTem
    End If
  Else
    MsgBox "No hay registros para agregar", vbCritical
  End If
End Sub

Private Sub CmdAgregarBandejaAnulacion_Click()
  Dim rstGuiasTemporal As New ADODB.Recordset
  AbrirRecorset rstGuiasTemporal, "SELECT Guia FROM temp_guias_despacho_anulado", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Do While rstGuiasTemporal.EOF = False
      AgregarGuia rstGuiasTemporal.Fields("Guia")
      rstGuiasTemporal.MoveNext
    Loop
  MsgBox rstGuiasTemporal.RecordCount & " Guias agregadas con exito", vbInformation
  CerrarRecorset rstGuiasTemporal
  AbrirRecorset rstUniversal, "TRUNCATE temp_guias_despacho_anulado", CnnPrincipal, adOpenDynamic, adLockOptimistic
End Sub

Private Sub CmdAgregarPorDocumento_Click()
  Dim rstGuiaDocumento As New ADODB.Recordset
  rstGuiaDocumento.CursorLocation = adUseClient
  Dim strSql As String
  If Principal.ToolConsultas1.AbrirDevDatos("Numero de documento", "Digite el numero del documento de la guia que desea agregar", 2, 0) = True Then
    strSql = "SELECT Guia FROM guias WHERE Despachada = 0 AND DocCliente = '" & Principal.ToolConsultas1.DatSt & "'"
    AbrirRecorset rstGuiaDocumento, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstGuiaDocumento.RecordCount > 0 Then
      AgregarGuia rstGuiaDocumento!Guia
    Else
      MsgBox "No se encontro la guia con ese numero de documento o no esta pendiente para despachar", vbCritical
    End If
    CerrarRecorset rstGuiaDocumento
    CmdAgregarPorDocumento_Click
  End If
End Sub

Private Sub CmdCargarPendientes_Click()
  If OptRuta.value = True Then
    If CboRuta.Tag <> "" Then
      AbrirRecorset rstTem, "SELECT guias.Guia, guias.CR, guias.IdCliente, guias.Cliente, guias.Remitente, guias.DocCliente, guias.NmDestinatario, ciudades.NmCiudad, guias.Unidades, guias.KilosReales, guias.KilosVolumen, guias.Estado, guias.IdDespacho, guias.COIng, guias.IdRuta From guias LEFT JOIN ciudades ON (guias.IdCiuDestino = ciudades.IdCiudad) Where (guias.Estado = 'I' OR guias.Estado = 'E') AND guias.CR = " & Coperaciones & " AND guias.Despachada=0 AND guias.IdRuta=" & Val(CboRuta.Tag) & " ORDER BY guias.FhEntradaBodega LIMIT " & Val(TxtNroResultados.Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      Set GrillaResultados.DataSource = rstTem
    Else
      MsgBox "No ha especificado una ruta para filtrar", vbCritical, "No ha especificado una ruta": CboRuta.SetFocus
    End If
  End If
  If OptCiudad.value = True Then
    If CboCiudad.Tag <> "" Then
      AbrirRecorset rstTem, "SELECT guias.Guia, guias.CR, guias.IdCliente, guias.Cliente, guias.Remitente, guias.DocCliente, guias.NmDestinatario, ciudades.NmCiudad, guias.Unidades, guias.KilosReales, guias.KilosVolumen, guias.Estado, guias.IdDespacho, guias.COIng, guias.IdRuta From guias LEFT JOIN ciudades ON (guias.IdCiuDestino = ciudades.IdCiudad) Where (guias.Estado = 'I' OR guias.Estado = 'E') AND guias.CR = " & Coperaciones & " AND guias.Despachada=0 AND guias.IdCiuDestino=" & Val(CboCiudad.Tag) & " ORDER BY guias.FhEntradaBodega LIMIT " & Val(TxtNroResultados.Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      Set GrillaResultados.DataSource = rstTem
    Else
      MsgBox "No ha especificado una ciudad para filtrar", vbCritical, "No ha especificado una ciudad": CboCiudad.SetFocus
    End If
  End If
  If OptTodos.value = True Then
    AbrirRecorset rstTem, "SELECT guias.Guia, guias.CR, guias.IdCliente, guias.Cliente, guias.Remitente, guias.DocCliente, guias.NmDestinatario, ciudades.NmCiudad, guias.Unidades, guias.KilosReales, guias.KilosVolumen, guias.Estado, guias.IdDespacho, guias.COIng, guias.IdRuta From guias LEFT JOIN ciudades ON (guias.IdCiuDestino = ciudades.IdCiudad) Where (guias.Estado = 'I' OR guias.Estado = 'E') AND guias.CR = " & Coperaciones & " AND guias.Despachada=0 ORDER BY guias.FhEntradaBodega LIMIT " & Val(TxtNroResultados.Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Set GrillaResultados.DataSource = rstTem
  End If
End Sub



Private Sub CmdPendientesPorImprimir_Click()
  FrmGuiasPorImprimir.Show 1
End Sub

Private Sub CmdQuitar_Click()
II = 1
Do While II <= LstRem.ListItems.Count
  If LstRem.ListItems(II).Checked = True Then
      TxtTotales(4).Text = Val(TxtTotales(4).Text) - Val(LstRem.ListItems(II).SubItems(4))
      TxtTotales(5).Text = Val(TxtTotales(5).Text) - Val(LstRem.ListItems(II).SubItems(5))
      TxtTotales(6).Text = Val(TxtTotales(6).Text) - 1
      TxtTotales(7).Text = Val(TxtTotales(7).Text) - Val(LstRem.ListItems(II).SubItems(6))
      TxtTotales(8).Text = Val(TxtTotales(8).Text) - Val(LstRem.ListItems(II).SubItems(9))
      AbrirRecorset rstUniversal, "SELECT VrFlete, VrManejo, Abonos, TipoCobro FROM guias WHERE Guia = " & Val(LstRem.ListItems(II)), CnnPrincipal, adOpenDynamic, adLockOptimistic
      If Val(rstUniversal!TipoCobro) = 2 Then
        TxtTotales(1).Text = Val(TxtTotales(1)) - ((Val(rstUniversal.Fields("VrFlete")) + Val(rstUniversal.Fields("VrManejo"))) - Val(rstUniversal.Fields("Abonos")))
      End If
      CerrarRecorset rstUniversal
      TxtTotales(2) = Val(TxtTotales(2)) - Val(LstRem.ListItems(II).SubItems(8))
      TxtTotales(3) = Val(TxtTotales(3)) - Val(LstRem.ListItems(II).SubItems(7))
      AbrirRecorset rstUniversal, "Update Guias Set IdDespacho=null, Estado='I', Despachada=0 Where Guia = " & LstRem.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
      LstRem.ListItems.Remove (II)
  Else
    II = II + 1
  End If
Loop
End Sub



Private Sub Form_Load()
Me.Caption = "Agregar y quitar guias del despacho... [" & FufuLo & "]"
Me.Tag = FufuLo
AbrirRecorset rstUniversal, "SELECT Guia, FhEntradaBodega, VrDeclarado, VrFlete, VrManejo, Abonos, Unidades, KilosReales, KilosVolumen, Estado, IdDespacho, Cliente, Recaudo, TipoCobro, Ciudades.NmCiudad, tipos_cobro.NmTipoCobro " & _
                            "FROM Guias " & _
                            "LEFT JOIN ciudades ON guias.IdCiuDestino = ciudades.IdCiudad " & _
                            "LEFT JOIN tipos_cobro ON guias.TipoCobro = tipos_cobro.IdTipoCobro " & _
                            "WHERE IdDespacho = " & Val(Me.Tag), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  IniProg 1, rstUniversal.RecordCount
  Do While rstUniversal.EOF = False
    Set Item = LstRem.ListItems.Add(, , rstUniversal!Guia)
    Item.SubItems(1) = Format(rstUniversal!FhEntradaBodega, "dd/mm")
    Item.SubItems(2) = rstUniversal!Cliente & ""
    Item.SubItems(3) = rstUniversal!NmCiudad & ""
    Item.SubItems(4) = rstUniversal!Unidades
    Item.SubItems(5) = rstUniversal!KilosReales
    Item.SubItems(6) = rstUniversal!KilosVolumen
    Item.SubItems(7) = rstUniversal!VrFlete
    Item.SubItems(8) = rstUniversal!VrManejo
    Item.SubItems(9) = rstUniversal!Recaudo
    Item.SubItems(10) = rstUniversal!NmTipoCobro & ""
    
    If Val(rstUniversal!TipoCobro) = 2 Then
      TxtTotales(1).Text = Val(TxtTotales(1)) + ((Val(rstUniversal.Fields("VrFlete")) + Val(rstUniversal.Fields("VrManejo"))) - Val(rstUniversal.Fields("Abonos")))
    End If
    TxtTotales(2).Text = Val(rstUniversal.Fields("VrManejo")) + Val(TxtTotales(2))
    TxtTotales(3).Text = Val(rstUniversal.Fields("VrFlete")) + Val(TxtTotales(3))
    TxtTotales(4).Text = rstUniversal.Fields("Unidades") + Val(TxtTotales(4))
    TxtTotales(5).Text = rstUniversal.Fields("KilosReales") + Val(TxtTotales(5))
    TxtTotales(6).Text = Val(TxtTotales(6)) + 1
    TxtTotales(7).Text = rstUniversal.Fields("KilosVolumen") + Val(TxtTotales(7))
    TxtTotales(8).Text = rstUniversal.Fields("Recaudo") + Val(TxtTotales(8))
    TxtTotales(9).Text = rstUniversal.Fields("VrDeclarado") + Val(TxtTotales(9))
    
    Prog (Val(TxtTotales(6).Text))
    rstUniversal.MoveNext
  Loop
  FinProg
CerrarRecorset rstUniversal
rstTem.CursorLocation = adUseClient
End Sub

Private Sub AgregarGuia(Guia As Long)
  Set Item = LstRem.FindItem(Guia)
  If Item Is Nothing Then
    AbrirRecorset rstUniversalSer, "SELECT Guia, FhEntradaBodega, Cliente, VrDeclarado, VrFlete, VrManejo, Abonos, Unidades, KilosReales, KilosVolumen, Estado, CR, IdDespacho, Recaudo, TipoCobro, Ciudades.NmCiudad, tipos_cobro.NmTipoCobro " & _
                                   "FROM guias " & _
                                   "LEFT JOIN ciudades ON guias.IdCiuDestino = ciudades.IdCiudad " & _
                                   "LEFT JOIN tipos_cobro ON guias.TipoCobro = tipos_cobro.IdTipoCobro " & _
                                   "WHERE Guia = " & Guia, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversalSer.EOF = False Then
      If rstUniversalSer!Estado = "A" Then
        MsgBox "Esta guia esta anulada, por lo tanto no se le puede agregar a un despacho", vbCritical, "Guia anulada"
      Else
        If rstUniversalSer!Estado <> "I" And rstUniversalSer!Estado <> "E" Then
          MsgBox "Solo puede agregarle guias a este despacho que esten impresas", vbCritical, "No se puede agregar"
        Else
          If rstUniversalSer!CR = Coperaciones Then
            If Val(rstUniversalSer!IdDespacho & "") = 0 Then
              Set Item = LstRem.ListItems.Add(, , rstUniversalSer!Guia)
                Item.SubItems(1) = Format(rstUniversalSer!FhEntradaBodega, "dd/mm")
                Item.SubItems(2) = rstUniversalSer!Cliente & ""
                Item.SubItems(3) = rstUniversalSer!NmCiudad & ""
                Item.SubItems(4) = rstUniversalSer!Unidades
                Item.SubItems(5) = rstUniversalSer!KilosReales
                Item.SubItems(6) = rstUniversalSer!KilosVolumen
                Item.SubItems(7) = rstUniversalSer!VrFlete
                Item.SubItems(8) = rstUniversalSer!VrManejo
                Item.SubItems(9) = rstUniversalSer!Recaudo
                Item.SubItems(10) = rstUniversalSer!NmTipoCobro & ""
      
              If rstUniversalSer!TipoCobro = 2 Then
                TxtTotales(1).Text = Val(TxtTotales(1)) + ((Val(rstUniversalSer.Fields("VrFlete")) + Val(rstUniversalSer.Fields("VrManejo"))) - Val(rstUniversalSer.Fields("Abonos")))
                'TxtTotales(0).Text = Val(TxtTotales(0)) + Val(rstUniversalSer.Fields("VrManejo"))
              End If
              TxtTotales(2).Text = Val(rstUniversalSer.Fields("VrManejo")) + Val(TxtTotales(2))
              TxtTotales(3).Text = Val(rstUniversalSer.Fields("VrFlete")) + Val(TxtTotales(3))
              TxtTotales(4).Text = rstUniversalSer.Fields("Unidades") + Val(TxtTotales(4))
              TxtTotales(5).Text = rstUniversalSer.Fields("KilosReales") + Val(TxtTotales(5))
              TxtTotales(6).Text = Val(TxtTotales(6)) + 1
              TxtTotales(7).Text = rstUniversalSer.Fields("KilosVolumen") + Val(TxtTotales(7))
              TxtTotales(8).Text = rstUniversalSer.Fields("Recaudo") + Val(TxtTotales(8))
              TxtTotales(9).Text = rstUniversalSer.Fields("VrDeclarado") + Val(TxtTotales(9))
              AbrirRecorset rstUniversal, "Update Guias Set IdDespacho=" & Val(Me.Tag) & ", Estado='P', Despachada=1 Where Guia = " & Guia, CnnPrincipal, adOpenDynamic, adLockOptimistic
            Else
              MsgBox "Esta guia ya esta en despacho Nro. [" & rstUniversalSer!IdDespacho & "] Si desea agregarla, tiene que sacarla del despacho al cual pertenece", vbInformation, "La guia pertenece a otro despacho"
            End If
          Else
            MsgBox "Esta guia esta en otro centro de operaciones, no la puede poner a viajar si no esta aqui", vbCritical
          End If
        End If
      End If
    Else
      MsgBox "La guia Nro. [" & Guia & "] NO existe verifique el numero", vbCritical, "La guia no existe"
    End If
    CerrarRecorset rstUniversalSer
  Else
    MsgBox "Esta guia ya fue agregada a este despacho", vbInformation, "La guia ya fue agregada"
  End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstTem = Nothing
End Sub

Private Sub LstRem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  LstRem.SortKey = ColumnHeader.Index - 1
  LstRem.Sorted = True
End Sub
