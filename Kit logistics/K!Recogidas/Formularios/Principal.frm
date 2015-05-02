VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8072DC64-8993-404F-8876-E5392C16A5C4}#1.0#0"; "PyConsultasKL.ocx"
Begin VB.MDIForm Principal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Recogidas [Kit Logistics]"
   ClientHeight    =   8415
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9660
   HelpContextID   =   3
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin ConsultasKL.ToolConsultas ToolConsultas1 
      Left            =   4560
      Top             =   3960
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.Timer TmPrincipal 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   840
      Top             =   120
   End
   Begin VB.PictureBox PicMensajes 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   9630
      TabIndex        =   0
      Top             =   8115
      Width           =   9660
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
      Left            =   120
      Top             =   120
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
            Picture         =   "Principal.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":19A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":267E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4032
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4D0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4FD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5138
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":529C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5B7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6856
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9560
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A23A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu MnuOpciones 
         Caption         =   "Opciones"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu MnuEdicion 
      Caption         =   "Edicion"
      Begin VB.Menu MnuBuscarAnuncios 
         Caption         =   "Buscar anuncios"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu MnuMantenimientoAnuncios 
      Caption         =   "Mantenimiento"
      Begin VB.Menu MnuAnuncios 
         Caption         =   "Ingresar Anuncios"
         Shortcut        =   ^I
      End
      Begin VB.Menu MnuRutasUrbanas 
         Caption         =   "Rutas Urbanas"
      End
      Begin VB.Menu MnuAuxiliares 
         Caption         =   "Auxliares"
      End
      Begin VB.Menu MnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuProgramarRecogidas 
         Caption         =   "Programar recogidas"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRecogidasPendientes 
         Caption         =   "Recogidas Pendientes"
      End
      Begin VB.Menu MnuRecogidasProgramadas 
         Caption         =   "Recogidas programadas"
      End
   End
   Begin VB.Menu MnuInformes 
      Caption         =   "Informes"
      Begin VB.Menu MnoVerRepRecogidas 
         Caption         =   "Todas las recogidas <fecha>"
      End
      Begin VB.Menu MnuRecogidasPorVehiculo 
         Caption         =   "Recogidas por vehiculo <Fecha>"
      End
      Begin VB.Menu MnuRecPendSinProg 
         Caption         =   "Recogidas pendientes <Sin programar>"
      End
      Begin VB.Menu MnuRecPendProg 
         Caption         =   "Recogidas pendientes <Programadas>"
      End
      Begin VB.Menu MnuAnalisisRutaRec 
         Caption         =   "Analisis de rutas <Recogidas>"
      End
      Begin VB.Menu MnuAnalisisRutaVeh 
         Caption         =   "Analisis de rutas <Vehiculos>"
      End
      Begin VB.Menu MnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCostosXRecogidaXCo 
         Caption         =   "Costos de recogidas"
      End
      Begin VB.Menu MnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfRutasUrbanas 
         Caption         =   "Rutas Urbanas"
      End
      Begin VB.Menu MnuInfAuxiliares 
         Caption         =   "Auxiliares"
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu MniIndiceAyuda 
         Caption         =   "Indice"
      End
      Begin VB.Menu MnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAcercade 
         Caption         =   "Acerca de"
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Load()
  Me.Caption = Me.Caption & "- CO [" & Coperaciones & "]"
  If GetSetting("Kit Logistics", "Recogidas", "Ini_Rec_Pend", 0) = 1 Then FrmRecogidasPendientes.Show 1
End Sub

Private Sub MnoVerRepRecogidas_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe de recogidas", 2) = True Then
    Mostrar_Reporte CnnPrincipal, 24, "select*from sql_ir_listadorecogidas where FhRecogida >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhRecogida<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'", "LISTADO DE RECOGIDAS POR FECHA", 2
  End If
End Sub

Private Sub MnuAcercade_Click()
  FrmAcercaDe.Show 1
End Sub
Private Sub MnuAnalisisRutaRec_Click()
  'If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite el rango de fechas para el cual desea ver el informe", 2) = True Then
   ' Mostrar_Reporte CnnPrincipal, 8, "Select*from SQL_IR_Analisis_Ruta_Recogidas where FhRecogida>='" & Principal.ToolConsultas1.Fecha1 & " 00:00:00" & "' and FhRecogida<='" & Principal.ToolConsultas1.Fecha2 & " 23:59:00" & "' and Coperaciones=" & Coperaciones, "", 2
    ' "Desde " & Principal.ToolConsultas1.Fecha1 & " Hasta " & Principal.ToolConsultas1.Fecha2
  'End If
End Sub
Private Sub MnuAnalisisRutaVeh_Click()
  'If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite el rango de fechas para el cual desea ver el informe", 2) = True Then
  '  Mostrar_Reporte CnnPrincipal, 9, "Select*from SQL_IR_Analisis_Ruta_Vehiculo where Fecha>='" & Principal.ToolConsultas1.Fecha1 & " 00:00:00" & "' and Fecha<='" & Principal.ToolConsultas1.Fecha2 & " 23:59:00" & "' and Coperaciones=" & Coperaciones, "", 2
    'CargarInformeRecogidas "Select*from SQL_IR_Analisis_Ruta_Vehiculo where Fecha>='" & Principal.ToolConsultas1.Fecha1 & " 00:00:00" & "' and Fecha<='" & Principal.ToolConsultas1.Fecha2 & " 23:59:00" & "' and Coperaciones=" & Coperaciones, "Desde " & Principal.ToolConsultas1.Fecha1 & " Hasta " & Principal.ToolConsultas1.Fecha2, 4, "Analisis de ruta <Vehiculos>"
  'End If
End Sub

Private Sub MnuAnuncios_Click()
  FrmAnunciosRecogida.Show
End Sub

Private Sub MnuAuxiliares_Click()
  FrmAuxiliares.Show 1
End Sub

Private Sub MnuBuscarAnuncios_Click()
  FrmBuscarAnuncios.Show 1
End Sub
Private Sub MnuCostosXRecogidaXCo_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Fecha Informe anuncios", "Digite el rango de fechas para el cual desea ver el informe", 2) = True Then
    'CargarInformeRecogidas "Select*from VehiculosRecogida where Fecha>='" & Principal.ToolConsultas1.Fecha1 & " 00:00:00" & "' and Fecha<='" & Principal.ToolConsultas1.Fecha2 & " 23:59:00" & "') and Coperaciones=" & Coperaciones, "Desde " & Principal.ToolConsultas1.Fecha1 & " Hasta " & Principal.ToolConsultas1.Fecha2, 2, "Informe de costos"
  End If
End Sub
Private Sub MnuInfAuxiliares_Click()
  Mostrar_Reporte CnnPrincipal, 7, "select*from SQL_IR_Auxiliares", "", 2
End Sub

Private Sub mnuInfRutasUrbanas_Click()
  Mostrar_Reporte CnnPrincipal, 6, "Select*from SQL_IR_RutasUrbanas", "", 2
End Sub

Private Sub MnuOpciones_Click()
  FrmOpciones.Show 1
End Sub

Private Sub MnuProgramarRecogidas_Click()
  FrmProgramarRecogidas.Show 1
End Sub

Private Sub MnuRecogidasPendientes_Click()
  FrmRecogidasPendientes.Show 1
End Sub

Private Sub MnuRecogidasPorVehiculo_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe de recogidas", 2) = True Then
    Mostrar_Reporte CnnPrincipal, 23, "select*from sql_ir_listadorecogidasvehiculo where Fecha >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and Fecha<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'", "RECOGIDAS DE LA ASIGNACION", 2
  End If
End Sub

Private Sub MnuRecogidasProgramadas_Click()
  FrmRecogidasProgramadas.Show 1
End Sub

Private Sub MnuRecPendProg_Click()
  Mostrar_Reporte CnnPrincipal, 2, "Select*from anuncios where efectiva=0 and Programada=1", "RECOGIDAS PENDIENTES Y PROGRAMADAS", 2
End Sub

Private Sub MnuRecPendSinProg_Click()
  Mostrar_Reporte CnnPrincipal, 2, "Select*from anuncios where efectiva=0 and Programada=0", "RECOGIDAS PENDIENTES SIN PROGRAMAR", 2
End Sub

Private Sub MnuRutasUrbanas_Click()
  FrmRutasUrbanas.Show 1
End Sub
Private Sub MnuSalir_Click()
  Unload Me
End Sub
Private Sub TmPrincipal_Timer()
  LblMensaje.Caption = ""
  TmPrincipal.Enabled = False
End Sub
