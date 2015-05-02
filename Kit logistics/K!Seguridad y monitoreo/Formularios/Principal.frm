VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8072DC64-8993-404F-8876-E5392C16A5C4}#1.0#0"; "PyConsultasKL.ocx"
Begin VB.MDIForm Principal 
   BackColor       =   &H8000000C&
   Caption         =   "Kit Logistics - Seguridad y monitoreo"
   ClientHeight    =   7950
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14490
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ConsultasKL.ToolConsultas ToolConsultas1 
      Left            =   3960
      Top             =   2760
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin MSComctlLib.ImageList IgListTool 
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
         NumListImages   =   15
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
            Picture         =   "Principal.frx":55E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5B80
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":611A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":66B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6C4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7528
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8202
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":AF0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":DC16
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
   Begin VB.Menu MnuMovimiento 
      Caption         =   "Movimiento"
      Begin VB.Menu MnuVehXIniMonitoreo 
         Caption         =   "Vehiculos por iniciar monitoreo"
      End
      Begin VB.Menu MnuVehiculosBetrieb 
         Caption         =   "Monitoreo"
      End
      Begin VB.Menu MnuHistoricoMonitoreos 
         Caption         =   "Historico Monitoreos"
      End
   End
   Begin VB.Menu MnuComplementos 
      Caption         =   "Complementos"
      Begin VB.Menu MnuPuestosControl 
         Caption         =   "Puestos de contros"
      End
   End
   Begin VB.Menu MnuReportes 
      Caption         =   "Reportes"
      Begin VB.Menu MnuVehiculosEnMonitoreo 
         Caption         =   "Lista monitoreos"
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MnuHistoricoMonitoreos_Click()
  FrmHistoricoMonitoreo.Show 1
End Sub

Private Sub MnuPuestosControl_Click()
  FrmPuestosControl.Show 1
End Sub

Private Sub MnuSalir_Click()
  Unload Me
End Sub

Private Sub MnuVehiculosBetrieb_Click()
  FrmBetrieb.Show
End Sub

Private Sub MnuVehiculosEnMonitoreo_Click()
    If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe", 2) = True Then
      Mostrar_Reporte CnnPrincipal, 50, "Select*from sql_mon_lista_monitoreos where FhHrSalida >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhHrSalida<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'", "", 2
    End If
End Sub

Private Sub MnuVehXIniMonitoreo_Click()
  FrmVehiculosPorMonitorear.Show 1
End Sub
