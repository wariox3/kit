VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{8072DC64-8993-404F-8876-E5392C16A5C4}#1.0#0"; "PyConsultasKL.ocx"
Begin VB.MDIForm Principal 
   BackColor       =   &H8000000C&
   Caption         =   "Principal"
   ClientHeight    =   6990
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11220
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ConsultasKL.ToolConsultas ToolConsultas1 
      Left            =   5160
      Top             =   1800
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.Timer TmPrincipal 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   4680
      Top             =   1800
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
      Top             =   6690
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
      Left            =   3600
      Top             =   1800
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
   Begin MSComDlg.CommonDialog CDExa 
      Left            =   4200
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu MnuConfiguracion 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu MnuProceso 
         Caption         =   "Proceso"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu MnuEdicion 
      Caption         =   "Edicion"
      Begin VB.Menu MnuBuscarFacturas 
         Caption         =   "Buscar Facturas"
      End
   End
   Begin VB.Menu MnuMantenimiento 
      Caption         =   "Facturacion"
      Begin VB.Menu MnuArchivoFacturas 
         Caption         =   "Facturas"
         Shortcut        =   ^F
      End
      Begin VB.Menu MnuPendientesPorFacturar 
         Caption         =   "Pendientes por facturar"
      End
   End
   Begin VB.Menu MenuCartera 
      Caption         =   "Cartera"
      Begin VB.Menu MnuCuentasCobrar 
         Caption         =   "Cuentas por cobrar"
      End
      Begin VB.Menu MnuRecibosCaja 
         Caption         =   "Recibos Caja"
         Shortcut        =   ^R
      End
      Begin VB.Menu MnuNotasCredito 
         Caption         =   "Notas credito"
      End
      Begin VB.Menu MnuNotasDebito 
         Caption         =   "Notas debito"
      End
   End
   Begin VB.Menu MnuContro 
      Caption         =   "Control"
      Begin VB.Menu MenuRecibosSinImprimir 
         Caption         =   "Recibos sin imprimir"
      End
   End
   Begin VB.Menu MnuProcesos 
      Caption         =   "Procesos"
      Begin VB.Menu MnuExportarContabilidad 
         Caption         =   "Exportar facturas contabilidad"
      End
      Begin VB.Menu MnuExportarRecibos 
         Caption         =   "Exportar recibos a contabilidad"
      End
      Begin VB.Menu MnuExportarNotasCredito 
         Caption         =   "Exportar notas credito"
      End
      Begin VB.Menu MnuControlFacturas 
         Caption         =   "Control facturas"
      End
      Begin VB.Menu MnuCorregirPuntoOperacionesFacturas 
         Caption         =   "Corregir punto operaciones"
      End
   End
   Begin VB.Menu MnuInformes 
      Caption         =   "Informes"
      Begin VB.Menu MnuCartera 
         Caption         =   "Cartera"
         Begin VB.Menu MnuCarteraEdades 
            Caption         =   "Cartera por edades (Cliente)"
         End
         Begin VB.Menu MnuCarteraEdadesAsesor 
            Caption         =   "Cartera por edades (Asesor)"
         End
         Begin VB.Menu MnuInfCarteraRecibos 
            Caption         =   "Recibos"
         End
         Begin VB.Menu MnuReporteNotasCredito 
            Caption         =   "Notas credito"
         End
      End
      Begin VB.Menu MnuPendientesFacturar 
         Caption         =   "Pendientes por facturar Corriente"
      End
      Begin VB.Menu MnuPendientesFacturarCli 
         Caption         =   "Pendientes por facturar (Cliente) Corriente"
      End
      Begin VB.Menu MnuListaFacturas 
         Caption         =   "Lista facturas"
      End
      Begin VB.Menu MnuFacturado 
         Caption         =   "Facturado"
      End
      Begin VB.Menu MnuFacturadoCli 
         Caption         =   "Facturado (Cliente)"
      End
      Begin VB.Menu MnuFacturadoAsesor 
         Caption         =   "Facturado (Asesor) detallado"
      End
      Begin VB.Menu MnuFacturasPorImprimir 
         Caption         =   "Facturas sin imprimir"
      End
      Begin VB.Menu MnuFacturacionConsolidado 
         Caption         =   "Facturacion consolidado"
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
  Me.Caption = Me.Caption & " V" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub MenuRecibosSinImprimir_Click()
  FrmRecibosSinImprimir.Show 1
End Sub

Private Sub MnuArchivoFacturas_Click()
  If CpPermiso(3, CodUsuarioActivo, 1, CnnPrincipal) = True Then
    FrmFacturas.Show
  End If
End Sub

Private Sub MnuBuscarFacturas_Click()
  FrmBuscarFacturas.Show 1
End Sub

Private Sub MnuCarteraPorEdades_Click()
  FrmInformeCarteraEdades.Show 1
End Sub

Private Sub MnuCarteraEdades_Click()
  FrmInformeCarteraEdades.Show 1
  If varParametrosCartera.Generar = True Then
    Mostrar_Reporte CnnPrincipal, 30, varParametrosCartera.sql, "Cartera por edades", 2
  End If
End Sub

Private Sub MnuCarteraEdadesAsesor_Click()
  FrmInformeCarteraEdades.Show 1
  If varParametrosCartera.Generar = True Then
    Mostrar_Reporte CnnPrincipal, 54, varParametrosCartera.sql, "Cartera por edades asesor", 2
  End If
End Sub

Private Sub MnuConfiguracion_Click()
  FrmConfiguracion.Show 1
End Sub
Private Sub MnuControlFacturas_Click()
  FrmControlFacturas.Show 1
End Sub

Private Sub MnuCorregirPuntoOperacionesFacturas_Click()
  Dim rstRecibos As New ADODB.Recordset
  rstRecibos.CursorLocation = adUseClient
  Dim rstRecibosDet As New ADODB.Recordset
  rstRecibosDet.CursorLocation = adUseClient
  Dim descuento As Double
  Dim ajustePeso As Double
  Dim retencionIca As Double
  Dim retencionFuente As Double
  Dim strSql As String
  rstRecibos.Open "SELECT recibos_caja.* FROM recibos_caja WHERE Fecha >= '2016-04-15'", CnnPrincipal, adOpenDynamic, adLockOptimistic
  MsgBox rstRecibos.RecordCount
  IniProg rstRecibos.RecordCount
  II = 1
  Do While rstRecibos.EOF = False
    descuento = 0
    ajustePeso = 0
    retencionIca = 0
    retencionFuente = 0
    strSql = "Select sum(descuento) as descuento, sum(ajuste_peso) as ajuste_peso, sum(retencion_ica) as retencion_ica, sum(retencion_fuente) as retencion_fuente from recibos_caja_det where IdRecibo=" & rstRecibos.Fields("IdRecibo")
    AbrirRecorset rstRecibosDet, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstRecibosDet.RecordCount > 0 Then
      descuento = Val(rstRecibosDet.Fields("descuento") & "")
      ajustePeso = Val(rstRecibosDet.Fields("ajuste_peso") & "")
      retencionIca = Val(rstRecibosDet.Fields("retencion_ica") & "")
      retencionFuente = Val(rstRecibosDet.Fields("retencion_fuente") & "")
      rstUniversal.Open "UPDATE recibos_caja SET descuento=" & descuento & ", ajuste_peso=" & ajustePeso & ", retencion_ica=" & retencionIca & ", retencion_fuente=" & retencionFuente & " WHERE IdRecibo = " & rstRecibos.Fields("IdRecibo"), CnnPrincipal, adOpenDynamic, adLockOptimistic
    End If
    CerrarRecorset rstRecibosDet
    rstRecibos.MoveNext
    II = II + 1
    Prog II
  Loop
  'rstUniversal.Open "UPDATE facturas_venta SET IdPO=1 WHERE IdPO is null", CnnPrincipal, adOpenDynamic, adLockOptimistic
End Sub

Private Sub MnuCuentasCobrar_Click()
  FrmCuentasCobrar.Show 1
End Sub

Private Sub MnuExportarContabilidad_Click()
  FrmExportarFacturas.Show 1
End Sub

Private Sub MnuExportarNotasCredito_Click()
  FrmExportarNotasCredito.Show 1
End Sub

Private Sub MnuExportarRecibos_Click()
  FrmExportarRecibos.Show 1
End Sub

Private Sub MnuFacturacionConsolidado_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe", 2) = True Then
    Mostrar_Reporte CnnPrincipal, 51, "Select*from sql_if_facturas_consolidado where (Fecha >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and Fecha<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00')", "Facturacion consolidado", 2
  End If
End Sub

Private Sub MnuFacturado_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe de lo que se ha facturado", 2) = True Then
    Mostrar_Reporte CnnPrincipal, 14, "Select*from sql_if_facturado where FhFac >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhFac<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'", "Facturado", 2
  End If
End Sub

Private Sub MnuFacturadoAsesor_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe", 2) = True Then
    Mostrar_Reporte CnnPrincipal, 49, "Select*from sql_if_facturado_asesor where (FhFac >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhFac<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00')", "Pendientes por facturar", 2
  End If
End Sub

Private Sub MnuFacturadoCli_Click()
  Principal.ToolConsultas1.AbrirDevConsulta 7, CnnPrincipal
  If Principal.ToolConsultas1.DatSt <> "" Then
    If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe de facturado por cliente", 2) = True Then
      Mostrar_Reporte CnnPrincipal, 14, "Select*from sql_if_facturado where IdCliente='" & Principal.ToolConsultas1.DatSt & "' and (FhFac >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhFac<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00')", "Pendientes por facturar", 2
    End If
  End If
End Sub



Private Sub MnuFacturasPorImprimir_Click()
  FrmFacturasPorImprimir.Show 1
End Sub

Private Sub MnuInformesFacturas_Click()

End Sub

Private Sub MnuInfCarteraRecibos_Click()
  FrmInformeRecibosCaja.Show 1
  If varParametrosRecibo.Generar = True Then
    If varParametrosRecibo.InformeDetallado = True Then
      Mostrar_Reporte CnnPrincipal, 57, varParametrosRecibo.sql, "Recibos detallados", 2
    Else
      Mostrar_Reporte CnnPrincipal, 56, varParametrosRecibo.sql, "Recibos", 2
    End If
  End If
End Sub

Private Sub MnuListaFacturas_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe", 2) = True Then
    Mostrar_Reporte CnnPrincipal, 47, "Select*from sql_if_lista_facturas where FhFac >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhFac<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'", "Listado facturas", 2
  End If
End Sub

Private Sub MnuNotasCredito_Click()
  FrmNotasCredito.Show 1
End Sub

Private Sub MnuNotasDebito_Click()
  FrmNotasDebito.Show 1
End Sub

Private Sub MnuPendientesFacturar_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe de pendientes por facturar", 2) = True Then
    Mostrar_Reporte CnnPrincipal, 10, "Select*from sql_if_pend_fac where FhEntradaBodega >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhEntradaBodega<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'", "Pendientes por facturar", 2
  End If
End Sub

Private Sub MnuPendientesFacturarCli_Click()
  Principal.ToolConsultas1.AbrirDevConsulta 7, CnnPrincipal
  If Principal.ToolConsultas1.DatSt <> "" Then
    If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe de pendientes por facturar", 2) = True Then
      Mostrar_Reporte CnnPrincipal, 10, "Select*from sql_if_pend_fac where Cuenta='" & Principal.ToolConsultas1.DatSt & "' and (FhEntradaBodega >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhEntradaBodega<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00')", "Pendientes por facturar", 2
    End If
  End If
End Sub

Private Sub MnuPendientesFacturarCon_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe de pendientes por facturar", 2) = True Then
    Mostrar_Reporte CnnPrincipal, 27, "Select*from sql_if_pend_fac_cont where FhEntradaBodega >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhEntradaBodega<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'", "Pendientes por facturar", 2
  End If
End Sub

Private Sub MnuPendientesFacturarDest_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe de pendientes por facturar", 2) = True Then
    Mostrar_Reporte CnnPrincipal, 28, "Select*from sql_if_pend_fac_dest where FhEntradaBodega >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhEntradaBodega<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'", "Pendientes por facturar", 2
  End If
End Sub

Private Sub MnuPendientesPorFacturar_Click()
  If CpPermisoEspecial(3, CodUsuarioActivo, CnnPrincipal) = True Then
    FrmPendientesPorFacturar.Show 1
  End If
End Sub

Private Sub MnuProceso_Click()
Dim rstActualizar As New ADODB.Recordset
rstActualizar.CursorLocation = adUseClient
'Dim NroFactura As Double

'AbrirRecorset rstUniversal, "SELECT IdFactura FROM facturas WHERE IdFactura <= 1", CnnPrincipal, adOpenDynamic, adLockOptimistic
'
'MsgBox rstUniversal.RecordCount
'IniProg rstUniversal.RecordCount
'II = 0
'Do While rstUniversal.EOF = False
'  NroFactura = 200000 + Val(rstUniversal.Fields("IdFactura"))
'  AbrirRecorset rstActualizar, "UPDATE facturas SET IdFactura = " & NroFactura & " WHERE IdFactura =" & rstUniversal.Fields("IdFactura"), CnnPrincipal, adOpenDynamic, adLockOptimistic
'  AbrirRecorset rstActualizar, "UPDATE guias SET IdFactura = " & NroFactura & " WHERE IdFactura =" & rstUniversal.Fields("IdFactura"), CnnPrincipal, adOpenDynamic, adLockOptimistic
'  AbrirRecorset rstActualizar, "UPDATE conceptosfacturas SET IdFactura = " & NroFactura & " WHERE IdFactura =" & rstUniversal.Fields("IdFactura"), CnnPrincipal, adOpenDynamic, adLockOptimistic
'  rstUniversal.MoveNext
'  II = II + 1
'  Prog II
'  DoEvents
'Loop
'MsgBox "Termino"
'FinProg
'CerrarRecorset rstUniversal
'AbrirRecorset rstUniversal, "SELECT IdFactura FROM facturas WHERE IdFactura >=1 AND IdFactura<=5000 AND Estado = 'A'", CnnPrincipal, adOpenDynamic, adLockOptimistic
'Do While rstUniversal.EOF = False
'  AbrirRecorset rstActualizar, "UPDATE facturas_venta SET Total = 0, VrFlete = 0, VrManejo = 0, VrOtros = 0 WHERE TipoFactura =1 AND Numero = " & rstUniversal.Fields("IdFactura"), CnnPrincipal, adOpenDynamic, adLockOptimistic
'  rstUniversal.MoveNext
'Loop
'CerrarRecorset rstUniversal
End Sub

Private Sub MnuRecibosCaja_Click()
  FrmReciboCaja.Show 1
End Sub

Private Sub MnuReporteNotasCredito_Click()
  FrmInformeNotaCredito.Show 1
  If varParametrosNotaCredito.Generar = True Then
    If varParametrosNotaCredito.InformeDetallado = True Then
      Mostrar_Reporte CnnPrincipal, 60, varParametrosNotaCredito.sql, "Nota credito detallados", 2
    Else
      Mostrar_Reporte CnnPrincipal, 59, varParametrosNotaCredito.sql, "Notas creditos", 2
    End If
  End If
End Sub

Private Sub MnuSalir_Click()
  Unload Me
End Sub



Private Sub TmPrincipal_Timer()
  LblMensaje.Caption = ""
  TmPrincipal.Enabled = False
End Sub
