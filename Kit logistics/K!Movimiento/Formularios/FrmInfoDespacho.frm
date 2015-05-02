VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmInfoDespacho 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informacion del despacho..."
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Aceptar / Salir"
      Height          =   255
      Left            =   6480
      TabIndex        =   36
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Frame FraEnc 
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   31
      Top             =   2040
      Width           =   2295
      Begin MSMask.MaskEdBox TxtFhCumplidos 
         Height          =   255
         Left            =   960
         TabIndex        =   32
         Top             =   525
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "dd-mmmm-yy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtFhExp 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   3
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   33
         Top             =   165
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         Format          =   "dd-mmmm-yy"
         PromptChar      =   "_"
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Cumplidos:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   525
         Width           =   765
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Expedida:"
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   34
         Top             =   165
         Width           =   705
      End
   End
   Begin VB.Frame FraVeCon 
      Enabled         =   0   'False
      Height          =   855
      Left            =   2520
      TabIndex        =   25
      Top             =   2040
      Width           =   6495
      Begin VB.CommandButton Command1 
         Caption         =   "Actualizar telefono Conductor"
         Height          =   255
         Left            =   4080
         TabIndex        =   40
         Top             =   160
         Width           =   2295
      End
      Begin VB.TextBox TxtTelConductor 
         Height          =   285
         Left            =   2280
         TabIndex        =   39
         Top             =   160
         Width           =   1695
      End
      Begin VB.TextBox TxtIdConductor 
         Height          =   285
         Left            =   960
         TabIndex        =   27
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtIdVehiculo 
         Height          =   285
         Left            =   960
         MaxLength       =   6
         TabIndex        =   26
         Top             =   160
         Width           =   1215
      End
      Begin VB.Label LblNmConductor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2280
         TabIndex        =   30
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Conductor:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vehiculo:"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   165
         Width           =   660
      End
   End
   Begin VB.Frame FraDatos 
      Caption         =   "Observaciones"
      Enabled         =   0   'False
      Height          =   975
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   8895
      Begin VB.TextBox TxtObservaciones 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   240
         Width           =   8655
      End
   End
   Begin VB.Frame FraResumen 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   2295
      Begin VB.CheckBox ChkLiquidado 
         Alignment       =   1  'Right Justify
         Caption         =   "Liquidada"
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
      Begin MSMask.MaskEdBox TxtUnidades 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   195
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPesoKilos 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   20
         Top             =   525
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Unidades:"
         Height          =   195
         Index           =   47
         Left            =   255
         TabIndex        =   22
         Top             =   200
         Width           =   720
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Peso kilos:"
         Height          =   195
         Index           =   49
         Left            =   210
         TabIndex        =   21
         Top             =   520
         Width           =   765
      End
   End
   Begin VB.Frame FraInfo 
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   8895
      Begin VB.TextBox TxtCO 
         Height          =   285
         Left            =   4920
         TabIndex        =   37
         Top             =   195
         Width           =   735
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "CO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   4560
         TabIndex        =   38
         Top             =   195
         Width           =   330
      End
      Begin VB.Label LblDespacho 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label LblManifiesto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3360
         TabIndex        =   15
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label LblUniversal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Despacho:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   51
         Left            =   120
         TabIndex        =   14
         Top             =   195
         Width           =   930
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Manifiesto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   2280
         TabIndex        =   13
         Top             =   195
         Width           =   945
      End
      Begin VB.Label LblUniversal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   5760
         TabIndex        =   12
         Top             =   195
         Width           =   660
      End
      Begin VB.Label LblEstado 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   6480
         TabIndex        =   11
         Top             =   195
         Width           =   2295
      End
   End
   Begin VB.Frame FraCit 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   6495
      Begin VB.TextBox TxtIdCiudadOrigen 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   200
         Width           =   1215
      End
      Begin VB.TextBox TxtNmCiuOrigen 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   200
         Width           =   4095
      End
      Begin VB.TextBox TxtNmCiuDestino 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   520
         Width           =   4095
      End
      Begin VB.TextBox TxtIdCiudadDestino 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   520
         Width           =   1215
      End
      Begin VB.TextBox TxtNmRuta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox TxtIdRuta 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Origen:"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   195
         Width           =   510
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Index           =   24
         Left            =   480
         TabIndex        =   8
         Top             =   855
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   285
         TabIndex        =   7
         Top             =   525
         Width           =   585
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   120
      X2              =   9000
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      Index           =   0
      X1              =   9000
      X2              =   120
      Y1              =   4080
      Y2              =   4080
   End
End
Attribute VB_Name = "FrmInfoDespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  AbrirRecorset rstUniversal, "SELECT*From Despachos where OrdDespacho=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.EOF = False Then
    LblManifiesto = rstUniversal!IdManifiesto
    LblDespacho = rstUniversal!OrdDespacho
    TxtFhExp = rstUniversal!FhExpedicion
    TxtFhCumplidos = rstUniversal!FhCumplidos & ""
    TxtIdRuta = rstUniversal!IdRuta
    TxtIdCiudadOrigen = rstUniversal!IdCiudadOrigen
    TxtIdCiudadDestino = rstUniversal!IdCiudadDestino
    TxtIdVehiculo = rstUniversal!IdVehiculo & ""
    TxtIdConductor = rstUniversal!IdConductor & ""
    TxtUnidades = rstUniversal!Unidades
    TxtPesoKilos = rstUniversal!KilosReales
    ChkLiquidado.Value = DevCheck(rstUniversal!liquidado)
    TxtCO = rstUniversal!CO
    LblEstado = DevEstadoDespacho(rstUniversal!Estado)
  End If
  If Val(TxtIdCiudadOrigen) <> 0 Then
    AbrirRecorset rstUniversal, "SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & TxtIdCiudadOrigen, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then TxtNmCiuOrigen = rstUniversal!NmCiudad & ""
    CerrarRecorset rstUniversal
  End If
  If Val(TxtIdCiudadDestino) <> 0 Then
    AbrirRecorset rstUniversal, "SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & TxtIdCiudadDestino, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then TxtNmCiuDestino = rstUniversal!NmCiudad & ""
    CerrarRecorset rstUniversal
  End If
  If Val(TxtIdRuta) <> 0 Then
    AbrirRecorset rstUniversal, "SELECT IdRuta, NmRuta From Rutas where IdRuta=" & TxtIdRuta, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then TxtNmRuta = rstUniversal!NmRuta & ""
    CerrarRecorset rstUniversal
  End If
  If TxtIdConductor.Text <> "" Then
    AbrirRecorset rstUniversal, "SELECT IdConductor, Nombre, Apellido1, Apellido2, Celular From Conductores Where IdConductor ='" & TxtIdConductor & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        LblNmConductor = rstUniversal!Nombre & " " & rstUniversal!Apellido1 & " " & rstUniversal!Apellido2 & ""
        TxtTelConductor.Text = rstUniversal!Celular
      End If
    CerrarRecorset rstUniversal
  End If
End Sub
