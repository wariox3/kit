VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmIntercambioEje 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Intercambio eje"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   14460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdImportar 
      Caption         =   "Importar"
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   7440
      Width           =   2535
   End
   Begin VB.CommandButton CmdVerificarLiquidacion 
      Caption         =   "Verificar liquidacion"
      Height          =   255
      Left            =   11760
      TabIndex        =   3
      Top             =   7440
      Width           =   2535
   End
   Begin MSDataGridLib.DataGrid GrillaGuias 
      Height          =   6855
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   12091
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "codigo_guia_pk"
         Caption         =   "ID"
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
         DataField       =   "numero_guia"
         Caption         =   "Numero"
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
         DataField       =   "forma_liquidacion"
         Caption         =   "F"
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
         DataField       =   "ct_unidades"
         Caption         =   "UND"
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
         DataField       =   "vr_declarado"
         Caption         =   "Declarado"
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
         DataField       =   "vr_flete"
         Caption         =   "Flete"
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
         DataField       =   "vr_manejo"
         Caption         =   "Manejo"
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
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdVerPendientes 
      Caption         =   "Ver pendientes importar"
      Height          =   255
      Left            =   5640
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo CboEntidades 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.CommandButton CmdCambiarCO 
      Caption         =   "Cambiar Centro de Operaciones"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   7440
      Width           =   3015
   End
   Begin VB.Frame FraCO 
      Caption         =   "Centros de operacion"
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   7440
      Width           =   4935
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   720
         TabIndex        =   8
         Tag             =   "1"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox TxtCOIng 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "C.O Ing:"
         Height          =   195
         Index           =   30
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   585
      End
   End
End
Attribute VB_Name = "FrmIntercambioEje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CnnEje As New ADODB.Connection
Public rstGuias As New ADODB.Recordset
Dim rstNegociacionKit As New ADODB.Recordset
Dim rstTerceroKit As New ADODB.Recordset

Private Sub CmdConectar_Click()

End Sub

Private Function LiquidarGuia(rstGuia As ADODB.Recordset) As Double
  Dim VrFlete As Double
  Dim TemListaPrecios As TipListaPrecios
  TemListaPrecios = DevListaPrecios(DevCodigoCiudadKit(rstGuia.Fields("codigo_ciudad_destino_fk")), DevCodigoProductoKit(rstGuia.Fields("codigo_producto_fk")), rstNegociacionKit.Fields("ListaPrecios"), 1, False)
  VrFlete = 0
  If TemListaPrecios.Devuelve = True Then
    Select Case Val(rstGuia.Fields("codigo_forma_liquidacion_fk"))
      Case 1
        VrFlete = rstGuia.Fields("ct_peso_liquidar") * TemListaPrecios.VrKilo
          
      Case 2
        VrFlete = rstGuia.Fields("ct_unidades") * TemListaPrecios.VrUnidad
      
      Case 3
        If (TemListaPrecios.KTope > rstGuia.Fields("ct_peso_liquidar")) Then
          Dim douPesoAdicional As Double
          douPesoAdicional = rstGuia.Fields("ct_peso_liquidar") - TemListaPrecios.KTope
          VrFlete = TemListaPrecios.VrKTope + (douPesoAdicional * TemListaPrecios.VrKdicional)
        Else
          VrFlete = TemListaPrecios.VrKTope
        End If
    End Select
    Print "Guia:" & rstGuia.Fields("codigo_guia_pk") & " Peso:" & rstGuia.Fields("ct_peso_liquidar") & " Flete:" & VrFlete
  End If
  LiquidarGuia = VrFlete
End Function
Private Function DevCodigoCiudadKit(codigoCiudadEje As Double) As String
  Dim rstCiudades As New ADODB.Recordset
  rstCiudades.CursorLocation = adUseClient
  AbrirRecorset rstCiudades, "SELECT codigo_interface_kit FROM gen_ciudades WHERE codigo_ciudad_pk = " & codigoCiudadEje, CnnEje, adOpenDynamic, adLockOptimistic
  DevCodigoCiudadKit = rstCiudades.Fields("codigo_interface_kit")
End Function
Private Function DevCodigoProductoKit(codigoProductoEje As Double) As String
  Dim rstProductos As New ADODB.Recordset
  rstProductos.CursorLocation = adUseClient
  AbrirRecorset rstProductos, "SELECT codigo_interface_kit FROM tte_productos WHERE codigo_producto_pk = " & codigoProductoEje, CnnEje, adOpenDynamic, adLockOptimistic
  DevCodigoProductoKit = rstProductos.Fields("codigo_interface_kit")
End Function
Private Function DevTipoCobroKit(codigoTipoCobroEje As Double) As String
  Dim rstGuiasTipos As New ADODB.Recordset
  rstGuiasTipos.CursorLocation = adUseClient
  AbrirRecorset rstGuiasTipos, "SELECT TipoCobro FROM guias_tipos WHERE IdGuiaTipo = " & codigoTipoCobroEje, CnnPrincipal, adOpenDynamic, adLockOptimistic
  DevTipoCobroKit = rstGuiasTipos.Fields("TipoCobro")
End Function
Private Function DevGuiFacKit(codigoTipoCobroEje As Double) As String
  Dim rstGuiasTipos As New ADODB.Recordset
  rstGuiasTipos.CursorLocation = adUseClient
  AbrirRecorset rstGuiasTipos, "SELECT GuiaFactura FROM guias_tipos WHERE IdGuiaTipo = " & codigoTipoCobroEje, CnnPrincipal, adOpenDynamic, adLockOptimistic
  DevGuiFacKit = rstGuiasTipos.Fields("GuiaFactura")
End Function

Private Function DevRuta(codigoCiudadDestino As Double) As String
  AbrirRecorset rstUniversal, "Select* from Rutas_Ciudades where IdCiudad=" & codigoCiudadDestino, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = False Then
    DevRuta = rstUniversal.Fields("IdRuta")
  Else
    DevRuta = 1
  End If
  CerrarRecorset rstUniversal
End Function
Private Function DevOrden(codigoCiudadDestino As Double) As String
  AbrirRecorset rstUniversal, "Select* from Rutas_Ciudades where IdCiudad=" & codigoCiudadDestino, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = False Then
    DevOrden = rstUniversal.Fields("Orden")
  Else
    DevOrden = 0
  End If
  CerrarRecorset rstUniversal
End Function
Private Sub CmdImportar_Click()
  Dim strSql As String
  Dim VrFlete As Double
  VrFlete = LiquidarGuia(rstGuias)
  If Val(rstGuias!numero_guia) <> 0 Then
    If ComprobarExGuiaGeneral(rstGuias!numero_guia) = False Then
      Dim codigoCiudadDestino As Double
      codigoCiudadDestino = DevCodigoCiudadKit(rstGuias!codigo_ciudad_destino_fk)
      strSql = "INSERT INTO Guias " & _
      "(Guia, CR, Remitente, IdCliente, DocCliente, NmDestinatario, DirDestinatario, TelDestinatario, IdCiuDestino, IdRuta, " & _
      "FhEntradaBodega, VrDeclarado, VrFlete, VrManejo, Unidades, KilosReales, KilosFacturados, KilosVolumen, " & _
      "Estado, IdFactura, IdDespacho, Observaciones, COIng, Cuenta, Cliente, Recaudo, orden, EmpaqueRef, RelCliente, IdCiuOrigen, TpServicio, CPorte, Entregada, Descargada, Despachada, Anulada, GuiFac, Facturada, IdUsuario, IdEmpresa, GuiaTipo, TipoCobro) " & _
      "VALUES(" & rstGuias!numero_guia & "," & Val(Campo(23).Text) & ",'" & rstTerceroKit!RazonSocial & "', '" & rstTerceroKit!IdCliente & "','" & rstGuias!documento_cliente & "','" & rstGuias!nombre_destinatario & "','" & rstGuias!direccion_destinatario & "','" & rstGuias!telefono_destinatario & "', " & codigoCiudadDestino & ", " & DevRuta(codigoCiudadDestino) & ", " & _
      "'2014-10-29', " & rstGuias!vr_declarado & ", " & VrFlete & ", " & rstGuias!vr_manejo & ", " & rstGuias!ct_unidades & ", " & rstGuias!ct_peso_real & ", " & rstGuias!ct_peso_liquidar & ", " & rstGuias!ct_peso_volumen & ", " & _
      "'D', 0, null, '" & rstGuias!Comentarios & "', " & Val(Campo(23).Text) & ", '" & rstTerceroKit!IdTercero & "', '" & rstTerceroKit!RazonSocial & "', " & rstGuias!vr_recaudo & ", " & DevOrden(codigoCiudadDestino) & ", '" & rstGuias!contenido & "','Rel cliente', 25332, 0, 0, 0, 0, 0, 0, " & DevGuiFacKit(rstGuias!codigo_tipo_pago_fk) & ", 0, 1, 1, " & rstGuias!codigo_tipo_pago_fk & ", " & DevTipoCobroKit(rstGuias!codigo_tipo_pago_fk) & ")"
      AbrirRecorset rstUniversal, strSql, CnnPrincipal, adOpenDynamic, adLockBatchOptimistic
    End If
  End If
End Sub

Private Sub CmdVerificarLiquidacion_Click()
  rstGuias.MoveFirst
  Do While rstGuias.EOF = False
    LiquidarGuia rstGuias
    rstGuias.MoveNext
  Loop
  rstGuias.MoveFirst
  
End Sub

Private Sub CmdVerPendientes_Click()
  Dim strSql As String
  strSql = "SELECT codigo_tercero_pk, nit FROM gen_terceros WHERE nombre_corto = '" & CboEntidades.Text & "'"
  AbrirRecorset rstUniversal, strSql, CnnEje, adOpenDynamic, adLockOptimistic
  If rstUniversal.RecordCount > 0 Then
    strSql = "SELECT tte_guias.* FROM tte_guias WHERE estado_despachada = 1 AND importado_kit_logistics = 0 and codigo_tercero_fk = " & rstUniversal!codigo_tercero_pk
    AbrirRecorset rstGuias, strSql, CnnEje, adOpenDynamic, adLockOptimistic
    Set GrillaGuias.DataSource = rstGuias
    
    AbrirRecorset rstTerceroKit, "SELECT IDTercero, IdCliente, RazonSocial FROM terceros WHERE IDTercero = '" & rstUniversal!nit & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
    AbrirRecorset rstNegociacionKit, "SELECT Id, ListaPrecios FROM negociaciones WHERE Id = " & rstTerceroKit.Fields("IdCliente"), CnnPrincipal, adOpenDynamic, adLockOptimistic
  End If
  CerrarRecorset rstUniversal
End Sub

Private Sub Form_Load()
  CnnEje.CursorLocation = adUseClient
  rstGuias.CursorLocation = adUseClient
  rstTerceroKit.CursorLocation = adUseClient
  rstNegociacionKit.CursorLocation = adUseClient
  Dim rstConfiguracion As New ADODB.Recordset
  rstConfiguracion.CursorLocation = adUseClient
  AbrirRecorset rstConfiguracion, "SELECT configuracion.* FROM configuracion", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Campo(23) = Coperaciones
  If Val(Campo(23).Text) <> 0 Then TxtCOIng.Text = DevResBus("SELECT IdPO, NmPuntoOperaciones From CentrosOperaciones where IdPO=" & Campo(23), "NmPuntoOperaciones", CnnPrincipal)
  On Error GoTo ErrCnn
    CnnEje.Open "DRIVER=" & rstConfiguracion!ejeDriver & "; SERVER=" & rstConfiguracion!ejeServidor & "; PORT=" & rstConfiguracion!ejePuerto & "; DATABASE=" & rstConfiguracion!ejeBaseDatos & "; PWD=" & rstConfiguracion!ejeClave & "; UID=" & rstConfiguracion!ejeUsuario & ";OPTION=3"
    Dim rstEntidades As New ADODB.Recordset
    rstEntidades.CursorLocation = adUseClient
    AbrirRecorset rstEntidades, "SELECT DISTINCT(codigo_tercero_fk), gen_terceros.nombre_corto FROM tte_guias LEFT JOIN gen_terceros ON tte_guias.codigo_tercero_fk = gen_terceros.codigo_tercero_pk WHERE importado_kit_logistics = 0", CnnEje, adOpenDynamic, adLockOptimistic
    CboEntidades.ListField = "nombre_corto"
    Set CboEntidades.RowSource = rstEntidades
    
ErrCnn:
    If Err.Number <> 0 Then
      MsgBox "No es posible conectar con esta cadena de conexion: " & Err.Description, vbCritical
      CmdVerPendientes.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set CnnEje = Nothing
  Set rstGuias = Nothing
  Set rstTerceroKit = Nothing
  Set rstNegociacionKit = Nothing
End Sub


