VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{8072DC64-8993-404F-8876-E5392C16A5C4}#1.0#0"; "PyConsultasKL.ocx"
Begin VB.MDIForm Principal 
   BackColor       =   &H8000000C&
   Caption         =   "Administracion de vehiculos..."
   ClientHeight    =   7395
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11220
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ConsultasKL.ToolConsultas ToolConsultas1 
      Left            =   5400
      Top             =   3480
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer TmPrincipal 
      Left            =   2040
      Top             =   600
   End
   Begin VB.PictureBox PicMensajes 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   11190
      TabIndex        =   0
      Top             =   7095
      Width           =   11220
      Begin MSComctlLib.ProgressBar PgsPrincipal 
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label LblTiutMensaje 
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   30
         Width           =   735
      End
      Begin VB.Label LblMensaje 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   30
         Width           =   9615
      End
   End
   Begin MSComctlLib.ImageList IgListTool 
      Left            =   2520
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":15A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":227E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":490C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4A70
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4BD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4E9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":577C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6456
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9160
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9E3A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu MnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu MnuAdministracion 
      Caption         =   "Administracion"
      Begin VB.Menu MnuVehiculos 
         Caption         =   "Vehiculos"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu MnuComplementos 
      Caption         =   "Complementos"
      Begin VB.Menu MnuComplementosArchivosBasicos 
         Caption         =   "Archivos Basicos"
         Begin VB.Menu MnuComArchBasicosCarrocerias 
            Caption         =   "Carrocerias"
         End
         Begin VB.Menu MnuComArchBasicosColores 
            Caption         =   "Colores"
         End
         Begin VB.Menu MnuComArchBasicosMarcas 
            Caption         =   "Marcas"
         End
         Begin VB.Menu MnuComArchBasicosLineas 
            Caption         =   "Lineas"
         End
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MnuComArchBasicosCarrocerias_Click()
  FrmCarrocerias.Show 1
End Sub

Private Sub MnuComArchBasicosColores_Click()
  FrmColores.Show 1
End Sub

Private Sub MnuComArchBasicosLineas_Click()
  FrmLineas.Show 1
End Sub

Private Sub MnuComArchBasicosMarcas_Click()
  FrmMarcas.Show 1
End Sub

Private Sub MnuSalir_Click()
  Unload Me
End Sub

Private Sub MnuVehiculos_Click()
  If CpPermiso(6, CodUsuarioActivo, 1, CnnPrincipal) = True Then
    FrmVehiculos.Show
  End If
End Sub
