VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8072DC64-8993-404F-8876-E5392C16A5C4}#1.0#0"; "PyConsultasKL.ocx"
Begin VB.MDIForm Principal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Movimiento [Kit Logistics - Logistica de transporte]  //"
   ClientHeight    =   7365
   ClientLeft      =   165
   ClientTop       =   570
   ClientWidth     =   11145
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ConsultasKL.ToolConsultas ToolConsultas1 
      Left            =   2400
      Top             =   1080
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.Timer TmMinuto 
      Interval        =   60000
      Left            =   2040
      Top             =   600
   End
   Begin MSComctlLib.ImageList IgListTool 
      Left            =   3000
      Top             =   600
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
   Begin MSComDlg.CommonDialog CDExa 
      Left            =   3600
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer TmPrincipal 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   2520
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
      ScaleWidth      =   11115
      TabIndex        =   0
      Top             =   7065
      Width           =   11145
      Begin MSComctlLib.ProgressBar PgsPrincipal 
         Height          =   255
         Left            =   900
         TabIndex        =   3
         Top             =   0
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   7440
         Top             =   0
         Width           =   375
      End
      Begin VB.Label LblMensaje 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   960
         TabIndex        =   1
         Top             =   30
         Width           =   9615
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
         TabIndex        =   2
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu MnuConfiguracion 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "Salir"
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu MnuBuscar 
      Caption         =   "Buscar"
      Begin VB.Menu MnuGuias 
         Caption         =   "Guias"
         Shortcut        =   ^B
      End
      Begin VB.Menu MnuBuscarViajes 
         Caption         =   "Viajes"
      End
   End
   Begin VB.Menu MnuMovimiento 
      Caption         =   "Guias"
      Begin VB.Menu mnuArchivoGuias 
         Caption         =   "Archivo"
         Shortcut        =   ^R
      End
      Begin VB.Menu MnuControlGuias 
         Caption         =   "Control guias"
      End
      Begin VB.Menu MnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuImprimirFormatoGuiaFactura 
         Caption         =   "Imprimir formato guia-factura"
      End
      Begin VB.Menu MnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu MnuActivarReDespacho 
         Caption         =   "Activar guia para Re-Despacho"
      End
      Begin VB.Menu MnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAgregarNovedad 
         Caption         =   "Agregar novedad"
      End
      Begin VB.Menu MnuSep21 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGuiasPorImprimir 
         Caption         =   "Guias por imprimir"
      End
      Begin VB.Menu MnuImpGuiaFormato 
         Caption         =   "Imprimir guia en formato"
      End
      Begin VB.Menu MnuSep 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu MnuCorregirGuia 
         Caption         =   "Corregir guia"
      End
      Begin VB.Menu MnuArchivoRelEntregaDoc 
         Caption         =   "Relaciones de cumplidos"
      End
      Begin VB.Menu MnuIntercambioEje 
         Caption         =   "Intercambio eje"
      End
   End
   Begin VB.Menu MnuDespachos 
      Caption         =   "Despachos"
      Begin VB.Menu MnuArchivoDespachos 
         Caption         =   "Archivo"
         Shortcut        =   ^D
      End
      Begin VB.Menu MnuSep22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDespachosPendientes 
         Caption         =   "Despachos pendientes"
      End
      Begin VB.Menu MnuSep23 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVerGuias 
         Caption         =   "Ver guias de un despacho"
      End
      Begin VB.Menu MnuSep24 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCorregirDespacho 
         Caption         =   "Corregir despacho"
      End
      Begin VB.Menu MnuCerrarDespacho 
         Caption         =   "Cerrar despacho"
         Shortcut        =   ^F
      End
      Begin VB.Menu MnuVerLiquidacion 
         Caption         =   "Liquidacion"
      End
      Begin VB.Menu MnuLiquidacion2 
         Caption         =   "Liquidacion 2"
      End
      Begin VB.Menu MnuResumenDespacho 
         Caption         =   "Resumen de despacho"
      End
   End
   Begin VB.Menu MnuAplicar 
      Caption         =   "A&plicar a despachos"
      Begin VB.Menu MnuEntregarGuias 
         Caption         =   "Cumplir entrega guias"
      End
      Begin VB.Menu MnuDescargar 
         Caption         =   "Descargar guias y despachos"
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCargarGuia 
         Caption         =   "Cargar guia"
      End
      Begin VB.Menu MnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDesembarcar 
         Caption         =   "Desembarcar"
      End
   End
   Begin VB.Menu MnuHerramientas 
      Caption         =   "Herramientas"
      Begin VB.Menu MnuMargenesRentabilidad 
         Caption         =   "Margenes de rentabilidad"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuAnalisisPendientesRuta 
         Caption         =   "Analisis de pendientes por Ruta"
      End
      Begin VB.Menu MnuAnalisisDespRuta 
         Caption         =   "Analisis de rentabilidad de despachos pendietes por ruta"
      End
      Begin VB.Menu MnuImportarGuias 
         Caption         =   "Importar Guias"
      End
      Begin VB.Menu MnuExportarDespachos 
         Caption         =   "Exportar despachos"
      End
      Begin VB.Menu MnuExportarDespachosContabilidad 
         Caption         =   "Exportar despachos contabilidad"
      End
      Begin VB.Menu MnuExportarGuiasFactura 
         Caption         =   "Exportar guias factura"
      End
      Begin VB.Menu MnuCuentasCobrar 
         Caption         =   "Cuentas por cobrar"
      End
      Begin VB.Menu MnuVerGuiasFactura 
         Caption         =   "Ver guias factura"
      End
      Begin VB.Menu MnuGenerarGuiasClientes 
         Caption         =   "Generar guias formatos"
      End
   End
   Begin VB.Menu MnuComplementos 
      Caption         =   "&Complementos"
      Begin VB.Menu MnuArcPrincipales 
         Caption         =   "Archivos Principales"
         Begin VB.Menu MnuTerceros 
            Caption         =   "Terceros"
         End
         Begin VB.Menu MnuClientes 
            Caption         =   "Negociaciones"
            Shortcut        =   ^T
         End
         Begin VB.Menu MnuArchivosPrincipalesConductores 
            Caption         =   "Conductores"
         End
         Begin VB.Menu MnuRemitentes 
            Caption         =   "Remitentes"
         End
         Begin VB.Menu MnuDestinatarios 
            Caption         =   "Destinatarios"
         End
         Begin VB.Menu MnuAsesoresComerciales 
            Caption         =   "Asesores"
         End
         Begin VB.Menu MnuCentroCostos 
            Caption         =   "Centros costos"
         End
      End
      Begin VB.Menu MnuArchivosSoporte 
         Caption         =   "Archivos soporte"
         Begin VB.Menu MnuArchivoSoporteCiudades 
            Caption         =   "Ciudades"
         End
         Begin VB.Menu MnuArchivosSoporteDepartamentos 
            Caption         =   "Departamentos"
         End
      End
      Begin VB.Menu MnuArchivosbasicos 
         Caption         =   "Archivos basicos"
         Begin VB.Menu MnuArchivosBasicosProductos 
            Caption         =   "Productos"
         End
         Begin VB.Menu MnuArchivosbasicosEmpaques 
            Caption         =   "Empaques"
         End
      End
      Begin VB.Menu MnuOrganizacion 
         Caption         =   "Organizacion"
         Begin VB.Menu MnuRutas 
            Caption         =   "Rutas"
         End
      End
   End
   Begin VB.Menu MnuRutinas 
      Caption         =   "Rutinas"
      Begin VB.Menu MnuAnalizarFrecuencias 
         Caption         =   "Analizar frecuencias"
      End
      Begin VB.Menu MnuDescargarDespCumplidos 
         Caption         =   "Descargar despachos cumplidos"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuInconsistenciasManejo 
         Caption         =   "Corregir inconsistencias"
      End
      Begin VB.Menu MnuCorregirDigitosVerificacion 
         Caption         =   "Corregir digitos de verificacion terceros"
      End
      Begin VB.Menu MnuCorregirCobrosDestino 
         Caption         =   "Corregir cobros al destino"
      End
      Begin VB.Menu MnuCorregirCodigosBarra 
         Caption         =   "Corregir codigos barra"
      End
   End
   Begin VB.Menu MnuInfRepLis 
      Caption         =   "Informes"
      Begin VB.Menu MnuPendientesPorEntregar 
         Caption         =   "Pendientes de entrega"
         Begin VB.Menu MnuPendRepTodo 
            Caption         =   "Guias pendientes por despacho"
         End
         Begin VB.Menu MnuPendDespachoCiudad 
            Caption         =   "Guias pendientes por despacho ciudad"
         End
         Begin VB.Menu MnuPendDevolucioones 
            Caption         =   "Guias pendientes por despacho <Devoluciones>"
         End
         Begin VB.Menu MnuPendCO 
            Caption         =   "Guias pendientes por despacho CO"
         End
      End
      Begin VB.Menu MnuListados 
         Caption         =   "Listados"
         Begin VB.Menu MnuListadosDatos 
            Caption         =   "Datos"
            Begin VB.Menu MnuListadoRemitentes 
               Caption         =   "Remitentes"
            End
            Begin VB.Menu MnuListadoClientes 
               Caption         =   "Clientes"
            End
            Begin VB.Menu MnuLestadoDestinatarios 
               Caption         =   "Destinatarios"
            End
            Begin VB.Menu MnuListadoTerceros 
               Caption         =   "Terceros"
            End
            Begin VB.Menu MnuListadoRutas 
               Caption         =   "Rutas"
            End
         End
         Begin VB.Menu MnuListadosControl 
            Caption         =   "Control"
            Begin VB.Menu MnuListadosControlAcuerdosCom 
               Caption         =   "Acuerdos Comerciales"
            End
         End
      End
      Begin VB.Menu MnuRepCaja 
         Caption         =   "Caja"
         Begin VB.Menu MnuVentasContado 
            Caption         =   "Ventas de contado"
         End
         Begin VB.Menu MnuGuiasFactura 
            Caption         =   "Guias factura"
         End
         Begin VB.Menu MnuRelacionRecibosCaja 
            Caption         =   "Relacion recibos caja"
         End
      End
      Begin VB.Menu MnuRepOperativos 
         Caption         =   "Operativos"
         Begin VB.Menu MnurepGuiasSinImprimir 
            Caption         =   "Guias sin imprimir"
            Begin VB.Menu MnurepGuiasSinImprimirTabla 
               Caption         =   "Registros"
            End
            Begin VB.Menu MnurepGuiasSinImprimirInf 
               Caption         =   "Reporte"
            End
         End
         Begin VB.Menu MnuRepDespaPendientes 
            Caption         =   "Despachos pendientes"
            Begin VB.Menu MnuDespPendInfRegistros 
               Caption         =   "Registros"
            End
            Begin VB.Menu MnuRepDespPenInf 
               Caption         =   "Reporte"
            End
         End
         Begin VB.Menu MnuContraentregasRecaudos 
            Caption         =   "Contraentregas y recaudos"
         End
         Begin VB.Menu MnuPendientesDescargar 
            Caption         =   "Pendientes por descargar"
         End
         Begin VB.Menu MnuListaDespachos 
            Caption         =   "Lista despachos"
         End
      End
      Begin VB.Menu MnuReportesGestion 
         Caption         =   "Gestion"
         Begin VB.Menu MnuNovedadesPendientes 
            Caption         =   "Novedades pendientes"
         End
         Begin VB.Menu MnuNovedadesPendientesCO 
            Caption         =   "Novedades pendientes centro operaciones"
         End
         Begin VB.Menu MnuTodasLasNovedades 
            Caption         =   "Todas las novedades"
         End
         Begin VB.Menu MnuInformeRedespachos 
            Caption         =   "Redespachos"
         End
         Begin VB.Menu MnuInformeResumenNovedades 
            Caption         =   "Resumen novedades"
         End
         Begin VB.Menu MnuInformeResumenNovedadesPend 
            Caption         =   "Resumen novedades pendientes"
         End
      End
      Begin VB.Menu MnuProduccionVentas 
         Caption         =   "Produccion, ventas"
         Begin VB.Menu MnuProduccionGeneral 
            Caption         =   "Produccion general"
         End
         Begin VB.Menu MnuProduccionPorCede 
            Caption         =   "Produccion por sede"
         End
         Begin VB.Menu MnuProduccionPorRuta 
            Caption         =   "Produccion por ruta"
         End
         Begin VB.Menu MnuProduccionPorAsesor 
            Caption         =   "Produccion por asesor"
         End
      End
      Begin VB.Menu MnuInfAdministrativos 
         Caption         =   "Administrativos"
         Begin VB.Menu MnuFletesCobradosVsFletesPagados 
            Caption         =   "Fletes Cobrados Vs Fletes Pagados"
         End
         Begin VB.Menu MnuDespachosSinCerrar 
            Caption         =   "Despachos sin cerrar"
         End
      End
      Begin VB.Menu MnuInformesControl 
         Caption         =   "Control"
         Begin VB.Menu MnuVencenDocumentos 
            Caption         =   "Documentos vencidos (Conductores y Vehiculos)"
         End
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu MnuContenidoAyuda 
         Caption         =   "Contenido"
      End
      Begin VB.Menu MnuIndiceAyuda 
         Caption         =   "Indice"
      End
      Begin VB.Menu MnuBusquedaAyuda 
         Caption         =   "Busqueda"
      End
      Begin VB.Menu MnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSoporteTecnico 
         Caption         =   "Soporte tecnico"
      End
      Begin VB.Menu MnuQueSoftwareWeb 
         Caption         =   "Que!Software en el web"
      End
      Begin VB.Menu MnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAcercaDe 
         Caption         =   "Acerca de Movimiento..."
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
  Dim Reg As Variant
  Me.Caption = "Kit Logistics " & App.Major & "." & App.Minor & "." & App.Revision & " - Transporte  //" & " Operaciones [" & Coperaciones & "]"
  MsgTit "BIENVENIDO AL KIT LOGISTICS FOR TRANSPOR"
  If GetSetting("Kit Logistics", "Movimiento", "Inicio_Analiis_Rutas") = 1 Then
    Me.Show
    FrmAnalisisPendientesPordDespachar.Show 1
  End If
End Sub

Private Sub MnuActivarReDespacho_Click()
    If Principal.ToolConsultas1.AbrirDevDatos("Digite el numero de la guia", "Digite el numero de la guia que desea activar para asignar aun nuevo viaje", 3, 0) = True Then
      FufuLo = Principal.ToolConsultas1.DatLo
      AbrirRecorset rstUniversal, "Select Guia, IdDespacho, Descargada,CR, COIng, Estado From Guias where Guia=" & FufuLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        If Val(rstUniversal!Descargada) = 0 Then
          If rstUniversal!Estado = "V" Or rstUniversal!Estado = "E" Or rstUniversal!Estado = "G" Or rstUniversal!Estado = "P" Then
            If rstUniversal!Estado = "P" Then
              Dim rstDespacho As New ADODB.Recordset
              rstDespacho.CursorLocation = adUseClient
              AbrirRecorset rstDespacho, "SELECT Estado FROM despachos WHERE OrdDespacho = " & rstUniversal.Fields("IdDespacho"), CnnPrincipal, adOpenDynamic, adLockOptimistic
              If rstDespacho.RecordCount > 0 Then
                If rstDespacho.Fields("Estado") = "V" Or rstDespacho.Fields("Estado") = "G" Then
                  AbrirRecorset rstUniversalAux, "Update Guias set Estado='I', IdDespacho=null, Despachada=0, Entregada=0, Descargada=0 where Guia=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
                  AbrirRecorset rstUniversal, "INSERT INTO Redespachos (Guia, Fecha, IdUsuario, IdDespacho) VALUES (" & FufuLo & ", '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "'," & CodUsuarioActivo & "," & Val(rstUniversal!IdDespacho) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
                  MsgBox "Guias activada para volver a Viajar, se creo un registro de Re-Viaje para soportar la accion", vbInformation, "Guia activada"
                End If
              End If
            Else
              AbrirRecorset rstUniversalAux, "Update Guias set Estado='I', IdDespacho=null, Despachada=0, Entregada=0, Descargada=0 where Guia=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
              AbrirRecorset rstUniversal, "INSERT INTO Redespachos (Guia, Fecha, IdUsuario, IdDespacho) VALUES (" & FufuLo & ", '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "'," & CodUsuarioActivo & "," & Val(rstUniversal!IdDespacho) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
              MsgBox "Guias activada para volver a Viajar, se creo un registro de Re-Viaje para soportar la accion", vbInformation, "Guia activada"
            End If
          Else
            MsgBox "Solo puede activar para un nuevo viaje las guias que esten con estado [VIAJANDO] o [DESEMBARCADA]", vbCritical, "La guia no esta viajando"
          End If
        Else
          MsgBox "No se pueden redespachar guias descargadas", vbCritical
        End If
        
      End If
      CerrarRecorset rstUniversal
    End If
End Sub




Private Sub MnuAgregarNovedad_Click()
    If Principal.ToolConsultas1.AbrirDevDatos("Digite la guia", "Digite la guia que desea procesar por novedad", 3, 0) = True Then
      FufuLo = Principal.ToolConsultas1.DatLo
      If ExRecorset("Select Guia from guias where Guia=" & FufuLo) = True Then
        II = 1
        FrmNovedades.Show 1
      Else
        MsgBox "la guia no existe", vbCritical
      End If
    End If
End Sub

Private Sub MnuAnalisisDespRuta_Click()
  FmrAnalisisDespachoRuta.Show 1
End Sub

Private Sub MnuAnalisisPendientesRuta_Click()
  FrmAnalisisPendientesPordDespachar.Show 1
End Sub

Private Sub MnuArchivoDespachos_Click()
    FrmManifiestos.Show 1
End Sub

Private Sub mnuArchivoGuias_Click()
  If CpPermiso(1, CodUsuarioActivo, 1, CnnPrincipal) = True Then
    FrmRemisiones.Show
  End If
End Sub

Private Sub MnuArchivoRelEntregaDoc_Click()
  FrmRelEntrega.Show 1
End Sub

Private Sub MnuArchivosbasicosEmpaques_Click()
  FrmEmpaques.Show 1
End Sub

Private Sub MnuArchivosBasicosProductos_Click()
  FrmProductos.Show 1
End Sub

Private Sub MnuArchivoSoporteCiudades_Click()
  FrmCiudades.Show 1
End Sub

Private Sub MnuArchivosPrincipalesConductores_Click()
  FrmConductores.Show 1
End Sub

Private Sub MnuArchivosSoporteDepartamentos_Click()
  FrmDepartamentos.Show 1
End Sub

Private Sub MnuAsesoresComerciales_Click()
  FrmAsesores.Show 1
End Sub

Private Sub MnuBuscarViajes_Click()
  FrmBuscarDespachosViaje.Show 1
End Sub

Private Sub MnuCargarGuia_Click()
    If Principal.ToolConsultas1.AbrirDevDatos("Digite el numero de la guia", "Digite el numero de la guia que desea volver a cargar y que quedara como pendiente", 3, 0) = True Then
      FufuLo = Principal.ToolConsultas1.DatLo
      AbrirRecorset rstUniversal, "Select Guia, IdDespacho, CR, COIng, Estado, Descargada From Guias where Guia=" & FufuLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        If Val(rstUniversal.Fields("Descargada")) = 1 Then
            AbrirRecorset rstUniversalAux, "Update Guias set Descargada=0 where Guia=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
            MsgBox "La guia se cargo correctamente y quedo como pendiente en el viaje", vbInformation, "Guia cargada con exito"
        Else
          MsgBox "La guia no esta descargada", vbCritical
        End If
        
      End If
      CerrarRecorset rstUniversal
    End If
End Sub

Private Sub MnuCentroCostos_Click()
  FrmCentrosCostos.Show 1
End Sub

Private Sub MnuCerrarDespacho_Click()
    If Principal.ToolConsultas1.AbrirDevDatos("Digite el numero del despacho", "Digite el numero del despacho que va a descargar", 3, 0) = True Then
      FufuLo = Principal.ToolConsultas1.DatLo
      AbrirRecorset rstUniversal, "Select OrdDespacho, IdManifiesto, Estado, Cerrado from Despachos Where OrdDespacho=" & FufuLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        If rstUniversal!Estado = "V" Or rstUniversal!Estado = "G" Then
          II = rstUniversal!Cerrado
          FrmCerrarDespacho.Show 1
        Else
          MsgBox "El estado del despacho es [" & DevEstadoDespacho(rstUniversal!Estado) & "] y solo se pueden cerrar los despachos con estado [VIAJANDO ó DESCARGADOS]", vbCritical, "Estado no válido"
        End If
      Else
        MsgBox "El despacho no existe", vbCritical, "El despacho no existe"
      End If
    End If
End Sub
Private Sub MnuClientes_Click()
  If CpPermiso(4, CodUsuarioActivo, 1, CnnPrincipal) = True Then
    FrmClientesNegociacion.Show 1
  End If
End Sub

Private Sub MnuConfiguracion_Click()
  'Dim rstGuia As New ADODB.Recordset
  'rstGuia.CursorLocation = adUseClient
  'Dim rstActualizar As New ADODB.Recordset
  'rstActualizar.CursorLocation = adUseClient
  'AbrirRecorset rstUniversal, "SELECT cuentas_cobrar.* from cuentas_cobrar where TipoFactura = 2 OR TipoFactura = 3", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  '  Do While rstUniversal.EOF = False
  '    FufuSt = "Select Guia, ciudades.NmCiudad from guias left join ciudades on guias.IdCiuDestino = ciudades.IdCiudad where Guia = " & rstUniversal!NroDocumento
  '    AbrirRecorset rstGuia, FufuSt, CnnPrincipal, adOpenDynamic, adLockOptimistic
  '    If rstGuia.RecordCount > 0 Then
  '      FufuSt = "update cuentas_cobrar set Soporte = '" & rstGuia!NmCiudad & "' where IdCxC = " & rstUniversal!IdCxC
  '      AbrirRecorset rstActualizar, FufuSt, CnnPrincipal, adOpenDynamic, adLockOptimistic
  '    End If
  '    CerrarRecorset rstGuia
  '    rstUniversal.MoveNext
  '  Loop
  'CerrarRecorset rstUniversal
  If CpPermisoEspecial(1, CodUsuarioActivo, CnnPrincipal) = True Then
    FrmConfiguracion.Show 1
  Else
    MsgBox "El usario no tiene permiso para ingresar a configuracion", vbCritical
  End If
End Sub

Private Sub MnuContenidoAyuda_Click()
   Dim hwndHelp As Long
   hwndHelp = HtmlHelp(hwnd, GetSetting("Kit Logistics", "Configuracion", "ArhivoAyuda"), &H0, 0)
End Sub

Private Sub MnuContraentregasRecaudos_Click()
  If Principal.ToolConsultas1.AbrirDevDatos("Contraentregas", "Ingrese el numero del despacho", 3, 0) = True Then
    Mostrar_Reporte CnnPrincipal, 8, "Select*from sql_im_contraentregas where iddespacho=" & Principal.ToolConsultas1.DatLo, "Contraentregas y recaudos", 2
  End If
End Sub

Private Sub MnuControlGuias_Click()
  FrmControlGuiasServicio.Show 1
End Sub

Private Sub MnuCorregirCobrosDestino_Click()
  'Dim strSql As String
  'Dim intNumeroDespacho As Integer
  'Dim rstPagos As New ADODB.Recordset
  'Dim rstDespachos As New ADODB.Recordset
  'Dim rstDespachoAct As New ADODB.Recordset
  'Dim rstGuias As New ADODB.Recordset
  'Dim douTotalCobroFleteDestino As Double
  'Dim douTotalCobroManejoDestino As Double
  'Dim douAbonos As Double
  'rstGuias.CursorLocation = adUseClient
  'rstDespachos.CursorLocation = adUseClient
  'rstPagos.CursorLocation = adUseClient
  'rstDespachoAct.CursorLocation = adUseClient
  'strSql = "SELECT OrdDespacho FROM despachos WHERE 1"
  'AbrirRecorset rstDespachos, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
  'Do While rstDespachos.EOF = False
  '  intNumeroDespacho = rstDespachos!OrdDespacho
  '  strSql = "SELECT Guia, VrFlete, VrManejo, Abonos FROM guias WHERE IdDespacho = " & intNumeroDespacho & " AND TipoCobro = 2"
  '  AbrirRecorset rstGuias, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
  '  douTotalCobroFleteDestino = 0
  '  douTotalCobroManejoDestino = 0
  '  douAbonos = 0
  '  Do While rstGuias.EOF = False
  '      douTotalCobroFleteDestino = douTotalCobroFleteDestino + Val(rstGuias!VrFlete)
  '      douTotalCobroManejoDestino = douTotalCobroManejoDestino + Val(rstGuias!VrManejo)
  '      If Val(rstGuias!Abonos) > 0 Then
  '        strSql = "SELECT VrFlete, VrManejo FROM recibos_caja_soporte WHERE Guia = " & rstGuias!Guia
  '        AbrirRecorset rstPagos, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
  '        Do While rstPagos.EOF = False
  '          douTotalCobroManejoDestino = douTotalCobroManejoDestino - Val(rstPagos!VrManejo)
  '          douTotalCobroFleteDestino = douTotalCobroFleteDestino - Val(rstPagos!VrFlete)
  '          rstPagos.MoveNext
  '        Loop
  '        CerrarRecorset rstPagos
  '        douAbonos = douAbonos + Val(rstGuias!Abonos)
  '      End If
  '    rstGuias.MoveNext
  '  Loop
  '  AbrirRecorset rstDespachoAct, "UPDATE Despachos SET ManejoCE = " & douTotalCobroManejoDestino & ", FleteCE = " & douTotalCobroFleteDestino & ", AbonosCE=" & douAbonos & ", TotalCE=" & douTotalCobroManejoDestino + douTotalCobroFleteDestino & " WHERE OrdDespacho = " & intNumeroDespacho, CnnPrincipal, adOpenDynamic, adLockOptimistic
  '  CerrarRecorset rstGuias
  '  strSql = "SELECT TipoCobro, SUM(VrFlete) as VrFlete, SUM(VrManejo) as VrManejo FROM guias WHERE IdDespacho = " & intNumeroDespacho & " GROUP BY TipoCobro"
  '  AbrirRecorset rstGuias, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
  '  Do While rstGuias.EOF = False
  '    If Val(rstGuias!TipoCobro) = 1 Then
  '      AbrirRecorset rstDespachoAct, "UPDATE Despachos SET FleteContado=" & rstGuias!VrFlete & ", ManejoContado=" & rstGuias!VrManejo & " WHERE OrdDespacho = " & intNumeroDespacho, CnnPrincipal, adOpenDynamic, adLockOptimistic
  '    End If
  '    If Val(rstGuias!TipoCobro) = 2 Then
  '      AbrirRecorset rstDespachoAct, "UPDATE Despachos SET FleteCETotal=" & rstGuias!VrFlete & ", ManejoCETotal=" & rstGuias!VrManejo & " WHERE OrdDespacho = " & intNumeroDespacho, CnnPrincipal, adOpenDynamic, adLockOptimistic
  '    End If
  '    If Val(rstGuias!TipoCobro) = 3 Then
  '      AbrirRecorset rstDespachoAct, "UPDATE Despachos SET FleteCorriente=" & rstGuias!VrFlete & ", ManejoCorriente=" & rstGuias!VrManejo & " WHERE OrdDespacho = " & intNumeroDespacho, CnnPrincipal, adOpenDynamic, adLockOptimistic
  '    End If
  '    rstGuias.MoveNext
  '  Loop
  '  CerrarRecorset rstGuias
  '  rstDespachos.MoveNext
  'Loop
End Sub

Private Sub MnuCorregirCodigosBarra_Click()
  AbrirRecorset rstUniversal, "Select guia from guias where CodigoBarras IS NULL", CnnPrincipal, adOpenDynamic, adLockOptimistic
  MsgBox "Este proceso tarda varios minutos por favor verificar que no este nadie en el sistema " & rstUniversal.RecordCount, vbCritical
  AbrirRecorset rstUniversal, "UPDATE guias SET CodigoBarras = concat('*',guia,'*') WHERE CodigoBarras IS NULL limit 50000 ", CnnPrincipal, adOpenDynamic, adLockOptimistic
End Sub

Private Sub MnuCorregirDespacho_Click()
  'If CpPermisoEspecial(10, CodUsuarioActivo, CnnPrincipal) = True Then
    If Principal.ToolConsultas1.AbrirDevDatos("Digite el despacho", "Digite el numero de despacho para corregir", 3, 0) = True Then
      FufuLo = Principal.ToolConsultas1.DatLo
      AbrirRecorset rstUniversal, "Select OrdDespacho from despachos where Estado != 'A' and OrdDespacho=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.RecordCount > 0 Then
        FrmCorregirDespacho.Show 1
      Else
        MsgBox "No se pueden corregir despachos: anulados", vbCritical
      End If
      CerrarRecorset rstUniversal
    End If
  'Else
  '  MsgBox "No tiene permiso para corregir guias", vbCritical
  'End If

End Sub

Private Sub MnuCorregirDigitosVerificacion_Click()
  FrmCorreccionDigitosVerificacion.Show 1
End Sub

Private Sub MnuCorregirGuia_Click()
  If CpPermisoEspecial(10, CodUsuarioActivo, CnnPrincipal) = True Then
    If Principal.ToolConsultas1.AbrirDevDatos("Digite la guia", "Digite la guia para corregir", 3, 0) = True Then
      FufuLo = Principal.ToolConsultas1.DatLo
      AbrirRecorset rstUniversal, "Select Guia, GuiFac from guias where Anulada=0 and Facturada=0 AND ExportadaCartera = 0 AND Guia=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.RecordCount > 0 Then
        If Val(rstUniversal.Fields("GuiFac")) = 1 Then
          If CpPermisoEspecial(14, CodUsuarioActivo, CnnPrincipal) = True Then
            FrmCorreccionGuia.Show 1
          Else
            MsgBox "La guia solicitada para correccion es factura de venta, debe tener un permiso especial para corregirla"
          End If
        Else
          FrmCorreccionGuia.Show 1
        End If
      Else
        MsgBox "No se pueden corregir guias: anuladas, facturadas, exportadas, tampoco las guias exportadas a contabilidad", vbCritical
      End If
      CerrarRecorset rstUniversal
    End If
  Else
    MsgBox "No tiene permiso para corregir guias", vbCritical
  End If
End Sub

Private Sub MnuCuentasCobrar_Click()
  FrmCuentasCobrar.Show 1
End Sub

Private Sub MnuDescargar_Click()
  If CpPermisoEspecial(2, CodUsuarioActivo, CnnPrincipal) = True Then
    If Principal.ToolConsultas1.AbrirDevDatos("Descargar...", "Ingrese el numero del despacho que desea descargar", 3, 0) = True Then
      FufuLo = Principal.ToolConsultas1.DatLo
      AbrirRecorset rstUniversalAux, "Select OrdDespacho, IdManifiesto, Estado from Despachos Where OrdDespacho=" & FufuLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversalAux.EOF = False Then
          Load FrmDescargar
          FrmDescargar.LblManifiesto.Caption = rstUniversalAux!IdManifiesto
          FrmDescargar.LblEstado = DevEstadoDespacho(rstUniversalAux!Estado)
          FrmDescargar.Show 1
      Else
        MsgBox "El despacho no existe", vbCritical, "El despacho no existe"
      End If
      CerrarRecorset rstUniversalAux
    End If
  End If
End Sub



Private Sub MnuDesembarcar_Click()
  Dim rstDespacho As New ADODB.Recordset
  rstDespacho.CursorLocation = adUseClient
  If Principal.ToolConsultas1.AbrirDevDatos("Desembarcar...", "Ingrese el numero del despacho que desea desembarcar", 3, 0) = True Then
    FufuLo = Principal.ToolConsultas1.DatLo
    AbrirRecorset rstDespacho, "Select OrdDespacho, IdManifiesto, Co, Estado from Despachos Where OrdDespacho=" & FufuLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstDespacho.EOF = False Then
      If rstDespacho!Estado = "V" Then
          Load FrmDesembarco
          FrmDesembarco.LblManifiesto.Caption = rstDespacho.Fields("IdManifiesto")
          FrmDesembarco.LblEstado.Caption = DevEstadoDespacho(rstDespacho!Estado)
          FrmDesembarco.Show 1
      Else
        MsgBox "El estado del despacho es [" & DevEstadoDespacho(rstDespacho!Estado) & "] y solo se pueden descargar los despachos con estado [VIAJANDO]", vbCritical, "Estado no válido"
      End If
    Else
      MsgBox "El despacho no existe", vbCritical, "El despacho no existe"
    End If
    CerrarRecorset rstUniversal
  End If
End Sub


Private Sub mnuDespachosPendientes_Click()
  FrmDespachosPendientes.Show 1
End Sub

Private Sub MnuDespachosSinCerrar_Click()
    Mostrar_Reporte CnnPrincipal, 25, "Select*from sql_im_despachosporcerrar ", "Despachos sin cerrar", 2
End Sub

Private Sub MnuDespPendInfRegistros_Click()
  FrmDespachosPendientes.Show 1
End Sub

Private Sub MnuDestinatarios_Click()
  FrmDestinatarios.Show 1
End Sub

Private Sub MnuEntregarGuias_Click()
  If Principal.ToolConsultas1.AbrirDevDatos("Entregar guias...", "Ingrese el numero del despacho que desea entregar guias", 3, 0) = True Then
    FufuLo = Principal.ToolConsultas1.DatLo
    AbrirRecorset rstUniversalAux, "Select OrdDespacho, IdManifiesto, Estado, FhExpedicion from despachos Where OrdDespacho=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstUniversalAux.EOF = False Then
        FrmEntregarGuias.LblFechaDespacho = rstUniversalAux.Fields("FhExpedicion")
        Load FrmEntregarGuias
        FrmEntregarGuias.Show 1
    Else
      MsgBox "El despacho no existe", vbCritical, "El despacho no existe"
    End If
    CerrarRecorset rstUniversalAux
  End If
End Sub


Private Sub MnuExportarDespachos_Click()
  FrmExportarManifiestos.Show 1
End Sub

Private Sub MnuExportarDespachosContabilidad_Click()
  FrmExportarManifiestosContabilidad.Show 1
End Sub

Private Sub MnuExportarGuiasFactura_Click()
  FrmExportarGuiasFactura.Show 1
End Sub

Private Sub MnuFletesCobradosVsFletesPagados_Click()
  Dim Criterio As String
  If CpPermisoEspecial(7, CodUsuarioActivo, CnnPrincipal) = True Then
    If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe de fletes", 2) = True Then
      If MsgBox("¿Desea filtrar por vehiculo?", vbYesNo + vbQuestion) = vbYes Then
        If Principal.ToolConsultas1.AbrirDevConsulta(5, CnnPrincipal) = True Then
          Criterio = " and IdVehiculo='" & Principal.ToolConsultas1.DatSt & "'"
        Else
          Criterio = ""
        End If
      End If
      Mostrar_Reporte CnnPrincipal, 20, "Select*from sql_im_flec_vs_flep where FhExpedicion >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhExpedicion<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'" & Criterio, "", 2
    End If
  End If
End Sub

Private Sub MnuGenerarGuiasClientes_Click()
  FrmGenerarGuiasClientes.Show 1
End Sub

Private Sub MnuGuias_Click()
  FrmBuscarGuias.Show 1
End Sub

Private Sub MnuGuiasFactura_Click()
  FrmGuiasFactura.Show 1
End Sub

Private Sub MnuGuiasPorImprimir_Click()
  FrmGuiasPorImprimir.Show 1
End Sub

Private Sub MnuImpGuiaFormato_Click()
  If Principal.ToolConsultas1.AbrirDevDatos("Numero de Guia", "Digite el numero de la guia que desea buscar", 3, 0) = True Then
    AbrirRecorset rstUniversal, "Select Guia from Guias where Guia=" & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      Mostrar_Reporte CnnPrincipal, 15, "Select*from sql_im_impguia where Guia=" & Principal.ToolConsultas1.DatLo, "", 2
      InsertarLog 7, Principal.ToolConsultas1.DatLo
    Else
      MsgBox "No se encontraron guias con este numero", vbCritical
    End If
    CerrarRecorset rstUniversal
  End If
End Sub

Private Sub MnuImportarGuias_Click()
  FrmImportarGuias.Show 1
End Sub

Private Sub MnuImportarGuiasArchivo_Click()

End Sub


Private Sub MnuImprimirFormatoGuiaFactura_Click()
  If Principal.ToolConsultas1.AbrirDevDatos("Digite el numero de la guia", "Digite el numero de la guia para ver el formato", 3, 0) = True Then
    Mostrar_Reporte CnnPrincipal, 33, "SELECT sql_im_formato_guia_factura.* FROM sql_im_formato_guia_factura WHERE Guia = " & Principal.ToolConsultas1.DatLo, "", 2
  End If
End Sub

Private Sub MnuInconsistenciasManejo_Click()
  AbrirRecorset rstUniversal, "update guias set VrFlete=0 where IdTpCtaFlete=4 or IdTpCtaFlete=5", CnnPrincipal, adOpenDynamic, adLockOptimistic
  AbrirRecorset rstUniversal, "update guias set VrManejo=0 where IdTpCtaManejo=4 or IdTpCtaManejo=5", CnnPrincipal, adOpenDynamic, adLockOptimistic
  MsgBox "Proceso terminado con exito", vbInformation
End Sub

Private Sub MnuInformeRedespachos_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe de redespachos", 2) = True Then
    Mostrar_Reporte CnnPrincipal, 37, "SELECT sql_im_redespachos.* FROM sql_im_redespachos WHERE Fecha >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and Fecha<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'", "", 2
  Else
    Mostrar_Reporte CnnPrincipal, 37, "SELECT sql_im_redespachos.* FROM sql_im_redespachos WHERE 1", "", 2
  End If
End Sub

Private Sub MnuInformeResumenNovedades_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe", 2) = True Then
    Mostrar_Reporte CnnPrincipal, 53, "Select * from sql_im_novedades where FhNovedad >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhNovedad<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'", "", 2
  End If
End Sub

Private Sub MnuInformeResumenNovedadesPend_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe", 2) = True Then
    Mostrar_Reporte CnnPrincipal, 53, "Select * from sql_im_novedades where Solucionada = 0 and FhNovedad >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhNovedad<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'", "", 2
  End If
End Sub

Private Sub MnuIntercambioEje_Click()
  FrmIntercambioEje.Show 1
End Sub

Private Sub MnuLiquidacion2_Click()
  If Principal.ToolConsultas1.AbrirDevDatos("Digite el numero del despacho", "Digite el numero del despacho que va a mostrar", 3, 0) = True Then
    FufuLo = Principal.ToolConsultas1.DatLo
    AbrirRecorset rstUniversal, "Select OrdDespacho from Despachos Where OrdDespacho=" & FufuLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      Mostrar_Reporte CnnPrincipal, 45, "Select*from sql_im_formato_liquidar_despacho2 where OrdDespacho=" & FufuLo, "Liquidacion despacho 2", 2
    Else
      MsgBox "El despacho no existe", vbCritical, "El despacho no existe"
    End If
  End If
End Sub

Private Sub MnuListaDespachos_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe", 2) = True Then
    Mostrar_Reporte CnnPrincipal, 48, "Select*from sql_im_lista_despachos where FhExpedicion >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhExpedicion<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'", "", 2
  End If
End Sub

Private Sub MnuListadoRutas_Click()
  Mostrar_Reporte CnnPrincipal, 1, "Select*from SQL_IM_rutas", "", 2
End Sub

Private Sub MnuListadosControlAcuerdosCom_Click()
'    Mostrar_Reporte CnnPrincipal, 2, "Select*from sql_im_cli_comercial", "", 2
End Sub

Private Sub MnuMargenesRentabilidad_Click()
  FrmMargenesRentabilidad.Show 1
End Sub

Private Sub MnuNovedadesPendientes_Click()
  Mostrar_Reporte CnnPrincipal, 9, "Select*from sql_im_novedadespendientes", "", 2
End Sub

Private Sub MnuNovedadesPendientesCO_Click()
  Mostrar_Reporte CnnPrincipal, 52, "Select*from sql_im_novedadespendientesco", "", 2
End Sub

Private Sub MnuPendCO_Click()
  FrmBuscarCO.Show 1
  If FufuLo <> 0 Then
    Mostrar_Reporte CnnPrincipal, 3, "Select*from sql_im_pend_desp where COIng=" & FufuLo, "", 2
  End If
End Sub

Private Sub MnuPendDespachoCiudad_Click()
  If CpPermisoEspecial(12, CodUsuarioActivo, CnnPrincipal) = True Then
    Mostrar_Reporte CnnPrincipal, 22, "Select*from sql_im_pend_desp where TpServicio<>5", "", 2
  End If
End Sub

Private Sub MnuPendDevolucioones_Click()
  Mostrar_Reporte CnnPrincipal, 3, "Select*from sql_im_pend_desp where TpServicio=5", "", 2
End Sub

Private Sub MnuPendientesDescargar_Click()
  Mostrar_Reporte CnnPrincipal, 12, "Select*from sql_im_pendescargar", "Pendientes por descargar", 2
End Sub

Private Sub MnuPendRepTodo_Click()
  If CpPermisoEspecial(12, CodUsuarioActivo, CnnPrincipal) = True Then
    If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe de pendiente por despachar", 2) = True Then
      Mostrar_Reporte CnnPrincipal, 3, "Select*from sql_im_pend_desp where TpServicio<>5 and FhEntradaBodega >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhEntradaBodega<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'", "", 2
    Else
      Mostrar_Reporte CnnPrincipal, 3, "Select*from sql_im_pend_desp where TpServicio<>5", "", 2
    End If
  End If
End Sub

Private Sub MnuProduccionGeneral_Click()
  If CpPermisoEspecial(4, CodUsuarioActivo, CnnPrincipal) = True Then
    If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe de produccion", 2) = True Then
      Mostrar_Reporte CnnPrincipal, 7, "Select*from slq_im_producciongral where FhEntradaBodega >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhEntradaBodega<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'", "", 2
    End If
  End If
End Sub

Private Sub MnuProduccionPorAsesor_Click()
  If CpPermisoEspecial(4, CodUsuarioActivo, CnnPrincipal) = True Then
    If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe de produccion", 2) = True Then
      Mostrar_Reporte CnnPrincipal, 42, "Select*from sql_im_produccionasesor where FhEntradaBodega >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhEntradaBodega<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'", "", 2
    End If
  End If
End Sub

Private Sub MnuProduccionPorCede_Click()
  If CpPermisoEspecial(4, CodUsuarioActivo, CnnPrincipal) = True Then
    FrmBuscarCO.Show 1
    If FufuLo <> 0 Then
      If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe de produccion", 2) = True Then
        Mostrar_Reporte CnnPrincipal, 7, "Select*from slq_im_producciongral where FhEntradaBodega >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhEntradaBodega<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00' and COIng=" & FufuLo, "", 2
      End If
    End If
  End If
End Sub

Private Sub MnuProduccionPorRuta_Click()
  If CpPermisoEspecial(4, CodUsuarioActivo, CnnPrincipal) = True Then
    If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe de produccion", 2) = True Then
      Mostrar_Reporte CnnPrincipal, 13, "Select*from sql_im_produccionruta where FhEntradaBodega >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FhEntradaBodega<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'", "", 2
    End If
  End If
End Sub

Private Sub MnuRelacionRecibosCaja_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para ver el informe de recibos de caaja", 2) = True Then
    Mostrar_Reporte CnnPrincipal, 46, "Select*from sql_im_recibos_caja where FechaRecibo >= '" & Format(Principal.ToolConsultas1.Fecha1, "yy-mm-dd") & " 00:00:00' and FechaRecibo<='" & Format(Principal.ToolConsultas1.Fecha2, "yy-mm-dd") & " 23:59:00'", "", 2
  End If
End Sub

Private Sub MnuRemitentes_Click()
  FrmRemitentes.Show
End Sub

Private Sub MnuRepDespPenInf_Click()
'    Mostrar_Reporte CnnPrincipal, 12, "Select*from SQL_IM_DespachosPendientesAbiertos", "", 2
End Sub

Private Sub MnurepGuiasSinImprimirInf_Click()
'  Mostrar_Reporte CnnPrincipal, 11, "Select*from SQL_IM_PendientesImprimir", "", 2
End Sub

Private Sub MnurepGuiasSinImprimirTabla_Click()
  FrmGuiasPorImprimir.Show 1
End Sub

Private Sub MnuResumenDespacho_Click()
  If Principal.ToolConsultas1.AbrirDevDatos("Digite el numero del despacho", "Digite el numero del despacho que va a mostrar", 3, 0) = True Then
    FufuLo = Principal.ToolConsultas1.DatLo
    AbrirRecorset rstUniversal, "Select OrdDespacho from Despachos Where OrdDespacho=" & FufuLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      Mostrar_Reporte CnnPrincipal, 19, "Select*from sql_im_resumendespacho where IdDespacho=" & FufuLo, "", 2
    Else
      MsgBox "El despacho no existe", vbCritical, "El despacho no existe"
    End If
  End If
End Sub

Private Sub MnuRutas_Click()
  FrmRutas.Show 1
End Sub
Private Sub MnuSalir_Click()
  Unload Me
End Sub

Private Sub MnuTerceros_Click()
  FrmTerceros.Show 1
End Sub

Private Sub MnuTodasLasNovedades_Click()
  FrmInformeNovedades.Show 1
End Sub

Private Sub MnuVencenDocumentos_Click()
  FrmAlertaDocumentosVencidos.Show 1
End Sub

Private Sub MnuVentasContado_Click()
  FrmVentasContado.Show 1
End Sub
Private Sub MnuVerGuias_Click()
  If CpPermisoEspecial(11, CodUsuarioActivo, CnnPrincipal) = True Then
    FrmVerGuiasDespacho.Show 1
  Else
    MsgBox "No tiene permisos para ver esta informacion", vbCritical
  End If
End Sub


Private Sub MnuVerGuiasFactura_Click()
  FrmGuiasFactura.Show 1
End Sub

Private Sub MnuVerLiquidacion_Click()
  If Principal.ToolConsultas1.AbrirDevDatos("Digite el numero del despacho", "Digite el numero del despacho que va a mostrar", 3, 0) = True Then
    FufuLo = Principal.ToolConsultas1.DatLo
    AbrirRecorset rstUniversal, "Select OrdDespacho from Despachos Where OrdDespacho=" & FufuLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      Mostrar_Reporte CnnPrincipal, 18, "Select*from slq_im_formatoliqdesp where OrdDespacho=" & FufuLo, "", 2
    Else
      MsgBox "El despacho no existe", vbCritical, "El despacho no existe"
    End If
  End If
End Sub

Private Sub TmMinuto_Timer()
  'On Error GoTo errMod
  'If Format(Time, "HH") = 8 Or Format(Time, "HH") = 11 Or Format(Time, "HH") = 14 Or Format(Time, "HH") = 16 Or Format(Time, "HH") = 19 Then
  '  FrmAlertaDocumentosVencidos.Show
  'End If
'errMod:
  
End Sub

Private Sub TmPrincipal_Timer()
  LblMensaje.Caption = ""
  TmPrincipal.Enabled = False
End Sub

