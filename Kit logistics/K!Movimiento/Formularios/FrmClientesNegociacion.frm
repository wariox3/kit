VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmClientesNegociacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Negociaciones / Acuerdos comerciales..."
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9390
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTInfo 
      Height          =   6375
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Informacion comercial"
      TabPicture(0)   =   "FrmClientesNegociacion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraOpciones"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraManejo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraCuentas"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FraMas"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FraDescuentos"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FraBas"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "FraDatos"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "FraServicios"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Recogidas"
      TabPicture(1)   =   "FrmClientesNegociacion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraRecogidas"
      Tab(1).ControlCount=   1
      Begin VB.Frame FraRecogidas 
         Caption         =   "Programacion de recogidas"
         Enabled         =   0   'False
         Height          =   2175
         Left            =   -74760
         TabIndex        =   46
         Top             =   480
         Width           =   7455
         Begin VB.TextBox TxtEncargadoRec 
            Height          =   285
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   51
            Top             =   1680
            Width           =   5415
         End
         Begin VB.TextBox TxtRuta 
            Height          =   285
            Left            =   1800
            MaxLength       =   5
            TabIndex        =   50
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox TxtDirRecogidas 
            Height          =   285
            Left            =   1800
            MaxLength       =   80
            TabIndex        =   49
            Top             =   960
            Width           =   5415
         End
         Begin VB.CheckBox ChKRecoge 
            Caption         =   "Recoger"
            Height          =   255
            Left            =   1800
            TabIndex        =   48
            Top             =   240
            Width           =   975
         End
         Begin MSComCtl2.DTPicker DPHoraRec 
            Height          =   255
            Left            =   1800
            TabIndex        =   47
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            Format          =   16711682
            UpDown          =   -1  'True
            CurrentDate     =   38971
         End
         Begin VB.Label LblCLientes 
            AutoSize        =   -1  'True
            Caption         =   "Encargado:"
            Height          =   195
            Index           =   7
            Left            =   840
            TabIndex        =   56
            Top             =   1680
            Width           =   825
         End
         Begin VB.Label LblConsulta 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   2
            Left            =   2520
            TabIndex        =   55
            Top             =   1320
            Width           =   4695
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ruta de recogidas:"
            Height          =   195
            Left            =   360
            TabIndex        =   54
            Top             =   1320
            Width           =   1350
         End
         Begin VB.Label LblCLientes 
            AutoSize        =   -1  'True
            Caption         =   "Direccion Recogidas:"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   53
            Top             =   960
            Width           =   1530
         End
         Begin VB.Label LblCLientes 
            AutoSize        =   -1  'True
            Caption         =   "Horario Recogidas:"
            Height          =   195
            Index           =   45
            Left            =   345
            TabIndex        =   52
            Top             =   615
            Width           =   1365
         End
      End
      Begin VB.Frame FraServicios 
         Caption         =   "Servicios"
         Enabled         =   0   'False
         Height          =   2415
         Left            =   6240
         TabIndex        =   44
         Top             =   1500
         Width           =   2535
         Begin MSComctlLib.ListView LstServicios 
            Height          =   2055
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   3625
            View            =   2
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame FraDatos 
         Enabled         =   0   'False
         Height          =   1020
         Left            =   120
         TabIndex        =   40
         Top             =   420
         Width           =   8655
         Begin VB.CheckBox ChkInactivo 
            Alignment       =   1  'Right Justify
            Caption         =   "Inactivo"
            Height          =   255
            Left            =   4800
            TabIndex        =   1
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TxtFechaIng 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7080
            TabIndex        =   57
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox TxtID 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   41
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox TxtNmNegociacion 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   0
            Top             =   600
            Width           =   7335
         End
         Begin VB.Label LblCLientes 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fecha Ingreso:"
            Height          =   195
            Index           =   1
            Left            =   6000
            TabIndex        =   58
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label LblCLientes 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ID:"
            Height          =   195
            Index           =   0
            Left            =   870
            TabIndex        =   43
            Top             =   240
            Width           =   210
         End
         Begin VB.Label LblCLientes 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Negociacion:"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   42
            Top             =   600
            Width           =   945
         End
      End
      Begin VB.Frame FraBas 
         Enabled         =   0   'False
         Height          =   2175
         Left            =   120
         TabIndex        =   36
         Top             =   4020
         Width           =   6015
         Begin VB.TextBox TxtObservaciones 
            BackColor       =   &H00FFFFFF&
            Height          =   1455
            Left            =   1200
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Text            =   "FrmClientesNegociacion.frx":0038
            Top             =   600
            Width           =   4605
         End
         Begin VB.TextBox TxtListaPrecioC 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   7
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox TxtNmListaPrecios 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   37
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label LblCLientes 
            AutoSize        =   -1  'True
            Caption         =   "Comentarios:"
            Height          =   195
            Index           =   42
            Left            =   150
            TabIndex        =   39
            Top             =   600
            Width           =   915
         End
         Begin VB.Label LblCLientes 
            AutoSize        =   -1  'True
            Caption         =   "Lista Precios:"
            Height          =   195
            Index           =   41
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.Frame FraDescuentos 
         Caption         =   "Liquidacion flete"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   2415
         Left            =   3480
         TabIndex        =   25
         Top             =   1500
         Width           =   2655
         Begin VB.CheckBox ChkRedondearFlete 
            Caption         =   "Redondear fletes"
            Height          =   375
            Left            =   120
            TabIndex        =   59
            Top             =   1560
            Width           =   2295
         End
         Begin VB.CheckBox ChkKilo 
            Alignment       =   1  'Right Justify
            Caption         =   "Kilo"
            Height          =   255
            Left            =   480
            TabIndex        =   31
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox ChkUnidad 
            Alignment       =   1  'Right Justify
            Caption         =   "Unidad"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox TxtMinimos 
            Height          =   285
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   3
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox TxtDctoKilo 
            Height          =   285
            Left            =   1200
            TabIndex        =   29
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox TxtDctoUni 
            Height          =   285
            Left            =   1200
            TabIndex        =   28
            Top             =   840
            Width           =   735
         End
         Begin VB.CheckBox ChkAdicional 
            Alignment       =   1  'Right Justify
            Caption         =   "Adicional"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1200
            Width           =   975
         End
         Begin VB.CheckBox ChkNoAplicaDctoReexpediciones 
            Caption         =   "Descuentos NO aplica para reexpediciones"
            Height          =   495
            Left            =   120
            TabIndex        =   26
            Top             =   1850
            Width           =   2295
         End
         Begin VB.Label LblCLientes 
            AutoSize        =   -1  'True
            Caption         =   "Dctos"
            Height          =   195
            Index           =   28
            Left            =   1200
            TabIndex        =   35
            Top             =   240
            Width           =   420
         End
         Begin VB.Label LblCLientes 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   32
            Left            =   2040
            TabIndex        =   34
            Top             =   840
            Width           =   120
         End
         Begin VB.Label LblCLientes 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   33
            Left            =   2040
            TabIndex        =   33
            Top             =   480
            Width           =   120
         End
         Begin VB.Label LblCLientes 
            AutoSize        =   -1  'True
            Caption         =   "Minimos:"
            Height          =   195
            Index           =   24
            Left            =   1200
            TabIndex        =   32
            Top             =   1200
            Width           =   615
         End
      End
      Begin VB.Frame FraMas 
         Caption         =   "Carta porte"
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         TabIndex        =   21
         Top             =   2340
         Width           =   3285
         Begin VB.OptionButton OptCp 
            Caption         =   "Opcional"
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   24
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptCp 
            Caption         =   "Si"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   23
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton OptCp 
            Caption         =   "No"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame FraCuentas 
         Caption         =   "Condiciones comerciales"
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         TabIndex        =   19
         Top             =   1500
         Width           =   3255
         Begin VB.TextBox TxtCupoCredito 
            Height          =   285
            Left            =   1200
            MaxLength       =   9
            TabIndex        =   2
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label LblCLientes 
            AutoSize        =   -1  'True
            Caption         =   "Cupo Credito:"
            Height          =   195
            Index           =   44
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   960
         End
      End
      Begin VB.Frame FraManejo 
         Caption         =   "Seguro o manejo:"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   1335
         Left            =   6240
         TabIndex        =   14
         Top             =   4020
         Width           =   2520
         Begin VB.TextBox TxtPorManejo 
            Height          =   285
            Left            =   840
            TabIndex        =   4
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox TxtVrUnidad 
            Height          =   285
            Left            =   840
            MaxLength       =   9
            TabIndex        =   5
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox TxtVrMinDespacho 
            Height          =   285
            Left            =   840
            TabIndex        =   6
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label LblCLientes 
            AutoSize        =   -1  'True
            Caption         =   "Min x Uni:"
            Height          =   195
            Index           =   37
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   705
         End
         Begin VB.Label LblCLientes 
            AutoSize        =   -1  'True
            Caption         =   "Manejo:"
            Height          =   195
            Index           =   35
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   570
         End
         Begin VB.Label LblCLientes 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   36
            Left            =   1440
            TabIndex        =   16
            Top             =   240
            Width           =   120
         End
         Begin VB.Label LblCLientes 
            AutoSize        =   -1  'True
            Caption         =   "Min Desp:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   720
         End
      End
      Begin VB.Frame FraOpciones 
         Enabled         =   0   'False
         Height          =   735
         Left            =   6240
         TabIndex        =   11
         Top             =   5460
         Width           =   2535
         Begin VB.CheckBox ChkPerListaGral 
            Caption         =   "Permitir lista de precios gral"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   2295
         End
         Begin VB.CheckBox ChkPerRecaudo 
            Caption         =   "Permitir recaudo"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   2295
         End
      End
   End
   Begin MSComctlLib.Toolbar ToolClientes 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   1005
      ButtonWidth     =   847
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
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
            Enabled         =   0   'False
            Caption         =   "Guar"
            Object.ToolTipText     =   "Guarda la informacio [F11]"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Object.ToolTipText     =   "Editar la informacion guardada [F10]"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Elim"
            Object.ToolTipText     =   "Elimina o anula el registro [F3]"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Can"
            Object.ToolTipText     =   "Cancela la creacion del registro [F4]"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Bus"
            Object.ToolTipText     =   "Buscar [Inicio]"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pri"
            Object.ToolTipText     =   "Ir al primer registro [F5]"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ant"
            Object.ToolTipText     =   "Ir al anterior registro [F6]"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sig"
            Object.ToolTipText     =   "Ir al siguiente registro [F7]"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ult"
            Object.ToolTipText     =   "Ir al ultimo registro [F8]"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cer"
            Object.ToolTipText     =   "Cerrar esta ventana [F12]"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Act"
            Object.ToolTipText     =   "Actualizar la informacion"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imp"
            Object.ToolTipText     =   "Imprimir registro [Fin]"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Car"
            Object.ToolTipText     =   "Carga informacion adicional [Pausa]"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Acc"
            Style           =   5
            Object.Width           =   1e-4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmClientesNegociacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Editando As Boolean
Dim rstClientes As New ADODB.Recordset

Private Sub ChkAdicional_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub ChkInactivo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub ChkKilo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub ChkPerListaGral_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub ChkPerRecaudo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub ChKRecoge_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub





Private Sub ChkUnidad_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub


Private Sub DPHoraRec_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  LaTecla (KeyCode), ToolClientes
End Sub
Private Sub Form_Load()
  IconosTool ToolClientes, Principal.IgListTool
  rstClientes.CursorLocation = adUseServer
  AbrirRecorset rstClientes, "SELECT*From Negociaciones", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Set Item = LstServicios.ListItems.Add(1, , "Paqueteo")
  Set Item = LstServicios.ListItems.Add(2, , "Semi-Masivo")
  Set Item = LstServicios.ListItems.Add(3, , "Masivo")
  Set Item = LstServicios.ListItems.Add(4, , "Urbanos/Local")
  Set Item = LstServicios.ListItems.Add(5, , "Encomiendas")
  Asignar rstClientes
  Editando = False
End Sub
Private Sub Asignar(rstAsignar As ADODB.Recordset)
  TxtId = rstAsignar!Id
  TxtNmNegociacion = rstAsignar!NmNegociacion & ""
  TxtFechaIng = rstAsignar!FecIng & ""
  ChkInactivo.value = DevCheck(rstAsignar.Fields("Inactivo"))
  TxtCupoCredito = rstAsignar!CupoCredito
  TxtPorManejo = rstAsignar!PorManejo
  TxtVrUnidad = rstAsignar!MinUniManejo
  TxtVrMinDespacho = rstAsignar!MinDesManejo
  ChkKilo.value = DevCheck(rstAsignar!ManKilo)
  TxtDctoKilo = rstAsignar!DctoK
  ChkUnidad.value = DevCheck(rstAsignar!ManUni)
  TxtDctoUni = rstAsignar!DctoU
  ChkAdicional.value = DevCheck(rstAsignar!ManAdicional)
  TxtMinimos = rstAsignar!Minimos
  DPHoraRec.value = rstAsignar!HorarioRecoge
  ChKRecoge.value = DevCheck(rstAsignar!Recoge)
  TxtDirRecogidas = rstAsignar!DirBodega & ""
  TxtRuta = rstAsignar!IdRutaRecogida & ""
  TxtEncargadoRec = rstAsignar!EncargadoRec & ""
  TxtListaPrecioC.Text = rstAsignar.Fields("ListaPrecios")
  TxtObservaciones = rstAsignar!Observaciones & ""
  OptCp(Val(rstAsignar!CartaPorte)).value = True
  LstServicios.ListItems(1).Checked = rstAsignar.Fields("ManPaqueteo").value
  LstServicios.ListItems(2).Checked = rstAsignar.Fields("ManSemiMasivo").value
  LstServicios.ListItems(3).Checked = rstAsignar.Fields("ManMasivo").value
  LstServicios.ListItems(4).Checked = rstAsignar.Fields("ManLocal").value
  LstServicios.ListItems(5).Checked = rstAsignar.Fields("ManEncomiendas").value
  ChkPerListaGral.value = DevCheck(rstAsignar!PermiteListaGral)
  ChkPerRecaudo.value = DevCheck(rstAsignar!PermiteRecaudo)
  ChkNoAplicaDctoReexpediciones.value = DevCheck(rstAsignar!NoAplicarDctoReexpediciones)
  ChkRedondearFlete.value = DevCheck(rstAsignar!RedondearFlete)
  LimpiarConsulta
End Sub
Private Sub limpiar()
  TxtId.Text = ""
  TxtFechaIng.Text = ""
  TxtNmNegociacion.Text = ""
  TxtCupoCredito.Text = ""
  TxtPorManejo.Text = ""
  TxtVrUnidad.Text = ""
  TxtVrMinDespacho.Text = ""
  TxtDctoKilo.Text = ""
  TxtDctoUni.Text = ""
  TxtMinimos.Text = ""
  DPHoraRec.value = Time
  TxtDirRecogidas.Text = ""
  TxtRuta.Text = ""
  TxtEncargadoRec.Text = ""
  TxtListaPrecioC.Text = ""
  TxtObservaciones.Text = ""
  ChKRecoge.value = 0
  ChkInactivo.value = 0
  ChkKilo.value = 0
  ChkUnidad.value = 0
  ChkAdicional.value = 0
  For II = 1 To LstServicios.ListItems.Count
    LstServicios.ListItems(II).Checked = False
  Next
  OptCp(2).value = True
  ChkPerRecaudo.value = 0
  ChkPerListaGral.value = 0
  ChkNoAplicaDctoReexpediciones.value = 0
  ChkRedondearFlete.value = 0
  LimpiarConsulta
End Sub
Private Sub LimpiarConsulta()
  LblConsulta(2).Caption = ""
  TxtNmListaPrecios.Text = ""
End Sub
Private Sub Desbloquear()
  BotTool 3, 17, ToolClientes, True
  FraDatos.Enabled = True
  
  FraCuentas.Enabled = True
  FraDescuentos.Enabled = True
  FraManejo.Enabled = True
  FraMas.Enabled = True
  FraBas.Enabled = True
  FraServicios.Enabled = True
  FraRecogidas.Enabled = True
  FraOpciones.Enabled = True
End Sub
Private Sub Bloquear()
  BotTool 3, 17, ToolClientes, False
  FraDatos.Enabled = False

  FraCuentas.Enabled = False
  FraDescuentos.Enabled = False
  FraManejo.Enabled = False
  FraMas.Enabled = False
  FraBas.Enabled = False
  FraServicios.Enabled = False
  FraRecogidas.Enabled = False
  FraOpciones.Enabled = False
End Sub
Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
      If CpPermiso(4, CodUsuarioActivo, 2, CnnPrincipal) = True Then
        Desbloquear
        limpiar
        TxtFechaIng.Text = Date
        TxtNmNegociacion.SetFocus
        SSTInfo.Tab = 0
      End If
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            Bloquear
            FufuSt = CStr(TxtListaPrecioC.Text)
            AbrirRecorset rstUniversal, "Update Negociaciones  set NmNegociacion='" & TxtNmNegociacion.Text & "', Minimos=" & Val(TxtMinimos) & ", MinUniManejo=" & Val(TxtVrUnidad) & ", MinDesManejo=" & Val(TxtVrMinDespacho) & ", ManKilo=" & ChkKilo.value & ", DctoK=" & Val(TxtDctoKilo) & ", ManUni=" & ChkUnidad.value & ", DctoU=" & Val(TxtDctoUni) & ", ManAdicional=" & ChkAdicional.value & ", PorManejo=" & Val(TxtPorManejo) & ", ListaPrecios=" & Val(TxtListaPrecioC.Text) & ", HorarioRecoge='" & Format(DPHoraRec.value, "h:m:s") & "', DirBodega='" & TxtDirRecogidas.Text & "', IdRutaRecogida=" & Val(TxtRuta.Text) & ", EncargadoRec='" & TxtEncargadoRec & "', Recoge=" & ChKRecoge.value & ", CupoCredito=" & Val(TxtCupoCredito) & ", Observaciones='" & TxtObservaciones.Text & "', " & _
            " CartaPorte=" & DevCartaPorte & ", Inactivo=" & ChkInactivo.value & ", ManPaqueteo=" & DevTpServicio(1) & ", ManSemiMasivo=" & DevTpServicio(2) & ", ManMasivo=" & DevTpServicio(3) & ", ManLocal=" & DevTpServicio(4) & ", ManEncomiendas=" & DevTpServicio(5) & ", PermiteRecaudo=" & ChkPerRecaudo.value & ", PermiteListaGral=" & ChkPerListaGral.value & ", NoAplicarDctoReexpediciones=" & ChkNoAplicaDctoReexpediciones.value & ", RedondearFlete=" & ChkRedondearFlete.value & " where ID=" & TxtId, CnnPrincipal, adOpenDynamic, adLockOptimistic
          End If
        Else
          AccionTool 7
        End If
        Editando = False
      Else
        If Validacion = True Then
          AbrirRecorset rstUniversal, "INSERT INTO Negociaciones (NmNegociacion, Minimos, MinUniManejo, MinDesManejo, ManKilo, DctoK, ManUni, DctoU, ManAdicional, PorManejo, ListaPrecios, HorarioRecoge, DirBodega, IdRutaRecogida, EncargadoRec, Recoge, CupoCredito, Observaciones, FecIng, CartaPorte, Inactivo, ManPaqueteo, ManSemiMasivo, ManMasivo, ManLocal, ManEncomiendas, PermiteRecaudo, PermiteListaGral, NoAplicarDctoReexpediciones, RedondearFlete) " & _
          "Values ('" & TxtNmNegociacion.Text & "', " & Val(TxtMinimos) & ", " & Val(TxtVrUnidad) & ", " & Val(TxtVrMinDespacho) & ", " & ChkKilo.value & ", " & Val(TxtDctoKilo) & ", " & ChkUnidad.value & ", " & Val(TxtDctoUni) & ", " & ChkAdicional.value & ", " & Val(TxtPorManejo) & ", " & Val(TxtListaPrecioC.Text) & ", '" & Format(DPHoraRec.value, "h:m:s") & "', '" & TxtDirRecogidas.Text & "', " & Val(TxtRuta.Text) & ", '" & TxtEncargadoRec & "', " & ChKRecoge.value & ", " & Val(TxtCupoCredito) & ", '" & TxtObservaciones.Text & "', '" & Format(TxtFechaIng.Text, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "'" & _
          ", " & DevCartaPorte & ", " & ChkInactivo.value & ", " & DevTpServicio(1) & ", " & DevTpServicio(2) & ", " & DevTpServicio(3) & ", " & DevTpServicio(4) & ", " & DevTpServicio(5) & ", " & ChkPerRecaudo.value & ", " & ChkPerListaGral.value & ", " & ChkNoAplicaDctoReexpediciones.value & ", " & ChkRedondearFlete.value & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
          Bloquear
          LimpiarConsulta
          AccionTool 17
          AccionTool 14
          Asignar rstClientes
        End If
      End If
    Case 5  'Editar
      If CpPermiso(4, CodUsuarioActivo, 3, CnnPrincipal) = True Then
        Editando = True
        Desbloquear
      End If
    Case 6 'Eliminar
      If CpPermiso(4, CodUsuarioActivo, 4, CnnPrincipal) = True Then
        MsgBox "No se pueden eliminar las negociaciones"
      End If
    Case 7 'Cancelar
      If Editando = True Then
        AccionTool 4
      Else
        Asignar rstClientes
        Bloquear
        LimpiarConsulta
      End If
    Case 9  'Buscar
      If Principal.ToolConsultas1.AbrirDevConsulta(2, CnnPrincipal) = True Then
        AbrirRecorset rstUniversal, "Select*from Negociaciones where Id=" & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          Asignar rstUniversal
        Else
          MsgBox "No se enconto el cliente, puede ser un error interno del sistema", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If
        
    Case 11 'Primero
      UPrimero rstClientes
      Asignar rstClientes
    Case 12 'Anterior
      UAnterior rstClientes
      Asignar rstClientes
    Case 13 'Siguiente
      USiguiente rstClientes
      Asignar rstClientes
    Case 14 'Ultimo
      UUltimo rstClientes
      Asignar rstClientes
    Case 16 'Cerrar
      CerrarRecorset rstClientes
      FufuSt = TxtId.Text
      Unload Me
    Case 17 'Actualizar
      rstClientes.Requery
    Case 18 'Imprimir
    Case 19
      If Val(TxtRuta.Text) <> 0 Then LblConsulta(2).Caption = DevResBus("SELECT IdRutaRec, NmRuta From RutasUrbanas where IdRutaRec=" & TxtRuta.Text, "NmRuta", CnnPrincipal)
      If Val(TxtListaPrecioC.Text) <> 0 Then TxtNmListaPrecios.Text = DevResBus("SELECT IdListaPrecios, NmListaPrecios From ListasPrecios where IdListaPrecios=" & TxtListaPrecioC, "NmListaPrecios", CnnPrincipal)
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FufuLo = Val(TxtId.Text)
End Sub

Private Sub LstServicios_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SSTInfo.Tab = 2
    SendKeys vbTab
  End If
End Sub
Private Sub OptCp_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub



Private Sub ToolClientes_ButtonClick(ByVal Button As MSComctlLib.Button)
  AccionTool Button.Index
End Sub
Private Function Validacion() As Boolean
  If Val(TxtCupoCredito) <> 0 Then
    If Val(TxtListaPrecioC) <> 0 Then
      Validacion = True
    Else
      MsgTit "El cliente debe tener una lista de precios para liquidar los fletes": Validacion = False: SSTInfo.Tab = 1: TxtListaPrecioC.SetFocus
    End If
  Else
    MsgTit "La cuenta (Cliente) debe tener un cupo de credito (Tope de facturacion)": Validacion = False: TxtCupoCredito.SetFocus: SSTInfo.Tab = 1
  End If
End Function
Private Function DevCartaPorte() As Byte
  If OptCp(0).value = True Then DevCartaPorte = 0
  If OptCp(1).value = True Then DevCartaPorte = 1
  If OptCp(2).value = True Then DevCartaPorte = 2
End Function

Private Sub TxtCupoCredito_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtCupoCredito, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtDctoKilo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtDctoKilo, KeyAscii, 3
End Sub
Private Sub TxtDctoUni_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtDctoUni, KeyAscii, 3
End Sub


Private Sub TxtDirRecogidas_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub







Private Sub TxtListaPrecioC_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsulta 11, CnnPrincipal
    TxtListaPrecioC.Text = Principal.ToolConsultas1.DatLo
  End If
End Sub

Private Sub TxtListaPrecioC_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtListaPrecioC, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtListaPrecioC_LostFocus()
  If Val(TxtListaPrecioC.Text) <> 0 Then
    AbrirRecorset rstUniversal, "SELECT IdListaPrecios, NmListaPrecios From ListasPrecios where IdListaPrecios=" & Val(TxtListaPrecioC), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      TxtNmListaPrecios.Text = rstUniversal!NmListaPrecios & ""
    Else
      TxtNmListaPrecios.Text = "": TxtListaPrecioC.Text = ""
    End If
  End If
End Sub

Private Sub TxtMinimos_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtMinimos, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub TxtNmNegociacion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtObservaciones_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys vbTab
  End If
End Sub

Private Sub TxtPorManejo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtPorManejo, KeyAscii, 3
End Sub
Private Sub TxtRuta_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirConsultaGral "IdRutaRec", "NmRuta", "RutasUrbanas", CnnPrincipal
    TxtRuta.Text = Principal.ToolConsultas1.DatLo
  End If
End Sub
Private Sub TxtRuta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtRuta, KeyAscii, 1
End Sub

Private Sub TxtRuta_Validate(Cancel As Boolean)
  If Val(TxtRuta) <> 0 Then
    AbrirRecorset rstUniversal, "SELECT IdRutaRec, NmRuta FROM RutasUrbanas where IdRutaRec=" & Val(TxtRuta), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      LblConsulta(2) = rstUniversal!NmRuta & ""
    Else
      LblConsulta(2) = "": TxtRuta = ""
    End If
    CerrarRecorset rstUniversal
  End If
End Sub

Private Sub TxtVrMinDespacho_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtPorManejo, KeyAscii, 3
End Sub

Private Sub TxtVrUnidad_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtVrUnidad, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Function DevTpServicio(Servicio As Byte) As Byte
  If LstServicios.ListItems(Servicio).Checked = True Then
    DevTpServicio = 1
  Else
    DevTpServicio = 0
  End If
End Function

