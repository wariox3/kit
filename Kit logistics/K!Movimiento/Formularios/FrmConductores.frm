VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmConductores 
   Caption         =   "Conductores"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   600
   End
   Begin VB.CommandButton CmdVer 
      Caption         =   "Ver"
      Height          =   255
      Left            =   9600
      TabIndex        =   59
      Top             =   2520
      Width           =   1935
   End
   Begin VB.PictureBox PicOrigen 
      Height          =   360
      Left            =   9600
      ScaleHeight     =   300
      ScaleWidth      =   585
      TabIndex        =   57
      Top             =   2880
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   9600
      TabIndex        =   56
      Top             =   720
      Width           =   1935
      Begin VB.PictureBox PicConductor 
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1275
         ScaleWidth      =   1635
         TabIndex        =   58
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame FraExternos 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   51
      Top             =   6480
      Width           =   3495
      Begin VB.TextBox TxtCampos 
         Enabled         =   0   'False
         Height          =   285
         Index           =   22
         Left            =   1080
         TabIndex        =   53
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha ingreso:"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Width           =   1050
      End
   End
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   5655
      Left            =   120
      TabIndex        =   24
      Top             =   720
      Width           =   9375
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   25
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   64
         Top             =   4800
         Width           =   2895
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   24
         Left            =   3480
         MaxLength       =   6
         TabIndex        =   7
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   23
         Left            =   4200
         TabIndex        =   60
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox ChkInactivo 
         Caption         =   "Inactivo"
         Height          =   255
         Left            =   1200
         TabIndex        =   55
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   7
         Left            =   6360
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   4
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox TxtNmCiudad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   50
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   3
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   0
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   2
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   1
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   17
         Left            =   1200
         MaxLength       =   248
         TabIndex        =   20
         Top             =   3720
         Width           =   7935
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   10
         Left            =   8520
         MaxLength       =   4
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   16
         Left            =   8640
         MaxLength       =   2
         TabIndex        =   17
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   15
         Left            =   6360
         MaxLength       =   50
         TabIndex        =   16
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   11
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   12
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   9
         Left            =   6360
         MaxLength       =   14
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   8
         Left            =   8520
         MaxLength       =   4
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   6
         Left            =   1200
         MaxLength       =   49
         TabIndex        =   5
         Top             =   2520
         Width           =   3495
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   18
         Left            =   6360
         MaxLength       =   30
         TabIndex        =   19
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   12
         Left            =   7200
         MaxLength       =   25
         TabIndex        =   13
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   20
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   22
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   13
         Left            =   6360
         TabIndex        =   14
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   14
         Left            =   6360
         MaxLength       =   50
         TabIndex        =   15
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   21
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   23
         Top             =   4440
         Width           =   5175
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   25
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   19
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   21
         Top             =   4080
         Width           =   7935
      End
      Begin MSComCtl2.DTPicker DTPFhNac 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   6
         Top             =   2880
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   50135043
         CurrentDate     =   37953
      End
      Begin MSComCtl2.DTPicker DTPVenceLic 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   6360
         TabIndex        =   18
         Top             =   2880
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   50135041
         CurrentDate     =   37953
      End
      Begin MSComCtl2.DTPicker DTPVenceSeguridadSocial 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   3960
         TabIndex        =   65
         Top             =   5160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   50135041
         CurrentDate     =   37953
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Vence seguridad social:"
         Height          =   195
         Index           =   25
         Left            =   1680
         TabIndex        =   66
         Top             =   5160
         Width           =   2190
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Seguridad Social:"
         Height          =   255
         Left            =   2295
         TabIndex        =   63
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Placa:"
         Height          =   195
         Left            =   3000
         TabIndex        =   62
         Top             =   3240
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   6
         Left            =   3840
         TabIndex        =   61
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ext:"
         Height          =   195
         Index           =   22
         Left            =   8160
         TabIndex        =   49
         Top             =   720
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Barrio:"
         Height          =   195
         Index           =   20
         Left            =   600
         TabIndex        =   48
         Top             =   2520
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Index           =   19
         Left            =   480
         TabIndex        =   47
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
         Height          =   195
         Index           =   18
         Left            =   360
         TabIndex        =   46
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Telefono 1:"
         Height          =   195
         Index           =   17
         Left            =   5505
         TabIndex        =   45
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ext:"
         Height          =   195
         Index           =   16
         Left            =   8160
         TabIndex        =   44
         Top             =   360
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Telefono 2:"
         Height          =   195
         Index           =   15
         Left            =   5505
         TabIndex        =   43
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Celular:"
         Height          =   195
         Index           =   14
         Left            =   5790
         TabIndex        =   42
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Licencia:"
         Height          =   195
         Index           =   13
         Left            =   5610
         TabIndex        =   41
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cat:"
         Height          =   195
         Index           =   12
         Left            =   8280
         TabIndex        =   40
         Top             =   2520
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fh Vence Lic:"
         Height          =   195
         Index           =   11
         Left            =   5280
         TabIndex        =   39
         Top             =   2880
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Señales:"
         Height          =   195
         Index           =   9
         Left            =   480
         TabIndex        =   38
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fec Nac:"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   37
         Top             =   2880
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         Height          =   195
         Index           =   5
         Left            =   855
         TabIndex        =   36
         Top             =   360
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   4
         Left            =   555
         TabIndex        =   35
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Alias:"
         Height          =   195
         Index           =   2
         Left            =   5880
         TabIndex        =   34
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Apellido 1:"
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   33
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Apellido 2:"
         Height          =   195
         Index           =   7
         Left            =   420
         TabIndex        =   32
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail:"
         Height          =   195
         Index           =   10
         Left            =   5835
         TabIndex        =   31
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Libreta:"
         Height          =   195
         Index           =   21
         Left            =   5790
         TabIndex        =   30
         Top             =   2160
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tel:"
         Height          =   195
         Index           =   23
         Left            =   840
         TabIndex        =   29
         Top             =   4440
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Otra comunicacion:"
         Height          =   195
         Index           =   24
         Left            =   5760
         TabIndex        =   28
         Top             =   1440
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ref Personal:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   4080
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
         Height          =   195
         Index           =   8
         Left            =   3150
         TabIndex        =   26
         Top             =   4440
         Width           =   720
      End
   End
   Begin MSComctlLib.Toolbar ToolConductores 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   11580
      _ExtentX        =   20426
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
Attribute VB_Name = "FrmConductores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstConductores As New ADODB.Recordset
Dim Editando As Boolean

Private Sub CmdVer_Click()
  FufuLo = TxtCampos(0).Text
  FrmVerImagen.Show 1
End Sub

Private Sub DTPFhNac_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub DTPVenceLic_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub DTPVenceSeguridadSocial_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  LaTecla (KeyCode), ToolConductores
End Sub
Private Sub Form_Load()
  IconosTool ToolConductores, Principal.IgListTool
  rstConductores.CursorLocation = adUseServer
  AbrirRecorset rstConductores, "SELECT conductores.* FROM conductores WHERE ConductorInactivo = 0", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Asignar rstConductores
End Sub

Private Sub Asignar(rstAsignar As ADODB.Recordset)
  For II = 0 To 25
    TxtCampos(II) = rstAsignar.Fields(II) & ""
  Next
  DTPVenceLic.value = rstAsignar.Fields("FhVenceLic")
  DTPFhNac.value = rstAsignar.Fields("FhNac")
  DTPVenceSeguridadSocial.value = rstAsignar.Fields("FhVenceSeguridadSocial")
  ChkInactivo.value = DevCheck(rstAsignar!ConductorInactivo)
  
  If CpExisteFichero(GetSetting("Kit Logistics", "Configuracion", "DirImagenes") & "conductores\" & TxtCampos(0).Text & ".jpg") = True Then
    PicOrigen = LoadPicture(GetSetting("Kit Logistics", "Configuracion", "DirImagenes") & "conductores\" & TxtCampos(0).Text & ".jpg")
    PicConductor.PaintPicture PicOrigen, 0, 0, PicConductor.ScaleWidth, PicConductor.ScaleHeight
  Else
    PicConductor.Picture = Nothing
  End If
End Sub

Private Sub limpiar()
  For II = 0 To 25
    TxtCampos(II).Text = ""
  Next
  TxtNmCiudad.Text = ""
End Sub
Private Sub Desbloquear()
  FraDatos.Enabled = True
  If CpPermisoEspecial(8, CodUsuarioActivo, CnnPrincipal) = False Then
    ChkInactivo.Enabled = False
  End If
  BotTool 3, 17, ToolConductores, True
End Sub
Private Sub Bloquear()
  FraDatos.Enabled = False
  ChkInactivo.Enabled = True
  BotTool 3, 17, ToolConductores, False
End Sub
Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
      If CpPermiso(5, CodUsuarioActivo, 2, CnnPrincipal) = True Then
        If Principal.ToolConsultas1.AbrirDevDatos("Digite ID del conductor", "Digite la identificacion del conductor (Cedula)", 2, 0) = True Then
          FufuSt = Principal.ToolConsultas1.DatSt
          If ExRecorset("Select IdConductor from Conductores where IdConductor='" & FufuSt & "'") = False Then
            Desbloquear
            limpiar
            TxtCampos(0).Text = FufuSt
            TxtCampos(1).SetFocus
            Editando = False
            DTPFhNac.value = Date
            DTPVenceLic.value = Date
            DTPVenceSeguridadSocial.value = Date
            TxtCampos(22).Text = Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s")
          Else
            MsgBox "Ya hay un conductor creado con esta identificacion, no se pueden crear dos conductores con esta identificacion", vbCritical, "El tercero ya existe"
          End If
        End If
      End If
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            Bloquear
            AbrirRecorset rstUniversal, "Update Conductores set Nombre='" & TxtCampos(1).Text & "', Apellido1='" & TxtCampos(2).Text & "', Apellido2='" & TxtCampos(3).Text & "', IdCiudad=" & Val(TxtCampos(4).Text) & ", Direccion='" & TxtCampos(5).Text & "', Barrio='" & TxtCampos(6).Text & "', TelConductor='" & TxtCampos(7).Text & "', Ext1='" & TxtCampos(8).Text & "', TelCOnductor2='" & TxtCampos(9).Text & "', Ext2='" & TxtCampos(10).Text & "', Celular='" & TxtCampos(11).Text & "', OtroCom='" & TxtCampos(12).Text & "', Email='" & TxtCampos(13).Text & "', Libreta='" & TxtCampos(14).Text & "', LicenciaConductor='" & TxtCampos(15).Text & "', Categoria='" & TxtCampos(16).Text & "', Senales='" & TxtCampos(17).Text & "', Alias='" & TxtCampos(18).Text & "', NmRefPersonal='" & TxtCampos(19).Text & "', TelRef='" & TxtCampos(20).Text & "', DirRef='" & TxtCampos(21).Text & "', " & _
            " FhVenceLic='" & Format(DTPVenceLic.value, "yyyy/mm/dd") & "', FhVenceSeguridadSocial='" & Format(DTPVenceSeguridadSocial.value, "yyyy/mm/dd") & "', TpIdConductor='" & TxtCampos(23).Text & "', PlacaPred='" & TxtCampos(24).Text & "', NroSeguridadSocial='" & TxtCampos(25).Text & "', FhNac='" & Format(DTPFhNac.value, "yyyy/mm/dd") & "', ConductorInactivo=" & ChkInactivo.value & " where IdConductor='" & TxtCampos(0).Text & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
            AccionTool 17
            Editando = False
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          AbrirRecorset rstUniversal, "INSERT INTO Conductores (IdConductor, Nombre, Apellido1, Apellido2, IdCiudad, Direccion, Barrio, TelConductor, Ext1, TelCOnductor2, Ext2, Celular, OtroCom, Email, Libreta, LicenciaConductor, Categoria, Senales, Alias, NmRefPersonal, TelRef, DirRef, FhIngreso, TpIdConductor, PlacaPred, NroSeguridadSocial, FhVenceLic, FhVenceSeguridadSocial, FhNac, ConductorInactivo)" & _
          " VALUES ('" & TxtCampos(0).Text & "', '" & TxtCampos(1).Text & "', '" & TxtCampos(2).Text & "', '" & TxtCampos(3).Text & "', " & Val(TxtCampos(4).Text) & ", '" & TxtCampos(5).Text & "', '" & TxtCampos(6).Text & "', '" & TxtCampos(7).Text & "', '" & TxtCampos(8).Text & "', '" & TxtCampos(9).Text & "', '" & TxtCampos(10).Text & "', '" & TxtCampos(11).Text & "', '" & TxtCampos(12).Text & "', '" & TxtCampos(13).Text & "', '" & TxtCampos(14).Text & "', '" & _
          TxtCampos(15).Text & "', '" & TxtCampos(16).Text & "', '" & TxtCampos(17).Text & "', '" & TxtCampos(18).Text & "', '" & TxtCampos(19).Text & "', '" & TxtCampos(20).Text & "', '" & TxtCampos(21).Text & "', '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', '" & TxtCampos(23).Text & "', '" & TxtCampos(24).Text & "', '" & TxtCampos(25).Text & "','" & Format(DTPVenceLic.value, "yyyy/mm/dd") & "','" & Format(DTPVenceSeguridadSocial.value, "yyyy/mm/dd") & "', '" & Format(DTPFhNac.value, "yyyy/mm/dd") & "', 0)", CnnPrincipal, adOpenDynamic, adLockOptimistic
          Bloquear
        End If
      End If
    Case 5  'Editar
      If CpPermiso(5, CodUsuarioActivo, 3, CnnPrincipal) = True Then
        Editando = True
        Desbloquear
      End If
    Case 6 'Eliminar
      If CpPermiso(5, CodUsuarioActivo, 4, CnnPrincipal) = True Then
        MsgBox "No se pueden eliminar conductores", vbCritical
      End If
    Case 7 'Cancelar
      If Editando = True Then
        AccionTool 4
      Else
        Asignar rstConductores
        Bloquear
      End If
    Case 9  'Buscar
      If Principal.ToolConsultas1.AbrirDevConsulta(3, CnnPrincipal) = True Then
        AbrirRecorset rstUniversal, "Select*from Conductores where IdConductor='" & Principal.ToolConsultas1.DatSt & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          Asignar rstUniversal
        Else
          MsgBox "No se encontraron conductores con este ID", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If
      
    Case 11 'Primero
      UPrimero rstConductores
      Asignar rstConductores
    Case 12 'Anterior
      UAnterior rstConductores
      Asignar rstConductores
    Case 13 'Siguiente
      USiguiente rstConductores
      Asignar rstConductores
    Case 14 'Ultimo
      UUltimo rstConductores
      Asignar rstConductores
    Case 16 'Cerrar
      CerrarRecorset rstConductores
      FufuSt = TxtCampos(0)
      Unload Me
    Case 17 'Actualizar
      rstConductores.Requery
    Case 18 'Imprimir
    Case 19
      If Val(TxtCampos(4).Text) <> 0 Then TxtNmCiudad.Text = DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & TxtCampos(4), "NmCiudad", CnnPrincipal)
      
  End Select
End Sub
Function Validacion() As Boolean
  Validacion = False
    If TxtCampos(1).Text <> "" Then
      If Val(TxtCampos(4).Text) <> 0 Then
        If TxtCampos(5).Text <> "" Then
          If TxtCampos(7).Text <> "" Then
            If TxtCampos(11).Text <> "" Then
              If TxtCampos(15).Text <> "" Then
                If TxtCampos(16).Text <> "" Then
                  If TxtCampos(23).Text <> "" Then
                    Validacion = True
                  Else
                    MsgBox "El conductor debe tener un tipo de identificacion", vbCritical: TxtCampos(23).SetFocus
                  End If
                Else
                  MsgBox "El conductor debe tener una categoria de licencia", vbCritical: TxtCampos(16).SetFocus
                End If
              Else
                MsgBox "El conductor debe tener una licencia", vbCritical: TxtCampos(15).SetFocus
              End If
            Else
              MsgBox "El conductor debe tener un celular", vbCritical: TxtCampos(11).SetFocus
            End If
          Else
            MsgBox "El conductor debe tener un telefono", vbCritical: TxtCampos(7).SetFocus
          End If
        Else
          MsgBox "El conductor debe tener una direccion", vbCritical: TxtCampos(5).SetFocus
        End If
      Else
        MsgBox "El conductor debe tener una ciudad", vbCritical: TxtCampos(4).SetFocus
      End If
    Else
      MsgBox "El conductor debe tener un nombre", vbCritical: TxtCampos(1).SetFocus
    End If
End Function

Private Sub Timer1_Timer()
  Asignar rstConductores
  Timer1.Enabled = False
End Sub

Private Sub ToolConductores_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub

Private Sub TxtCampos_GotFocus(Index As Integer)
  EnfocarT TxtCampos(Index)
  TxtCampos(Index).BackColor = &H80000001
  TxtCampos(Index).ForeColor = &HFFFFFF
End Sub

Private Sub TxtCampos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Select Case Index
      Case 4
        Principal.ToolConsultas1.AbrirDevConsulta 1, CnnPrincipal
        TxtCampos(4).Text = Principal.ToolConsultas1.DatLo
    End Select
  End If
End Sub

Private Sub TxtCampos_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  Select Case Index
    Case 4, 7, 8, 9, 10, 11, 12, 15, 20
      ValidarEntrada TxtCampos(1), KeyAscii, 1
  End Select
End Sub
Private Sub TxtCampos_LostFocus(Index As Integer)
  TxtCampos(Index).BackColor = &H80000005
  TxtCampos(Index).ForeColor = &H80000012
End Sub

Private Sub TxtCampos_Validate(Index As Integer, Cancel As Boolean)
  Select Case Index
    Case 4
      If Val(TxtCampos(4).Text) <> 0 Then
        AbrirRecorset rstUniversal, "SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & TxtCampos(4), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          TxtNmCiudad.Text = rstUniversal!NmCiudad & ""
        Else
          TxtNmCiudad.Text = "": TxtCampos(4).Text = ""
        End If
        CerrarRecorset rstUniversal
      End If
  End Select
End Sub



