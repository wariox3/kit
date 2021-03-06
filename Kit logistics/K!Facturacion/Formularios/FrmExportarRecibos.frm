VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmExportarRecibos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar recibos caja"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   11340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExportarContai 
      Caption         =   "Exportar Contai"
      Height          =   375
      Left            =   7920
      TabIndex        =   11
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton CmdSeleccionarTodo 
      Caption         =   "Seleccionar todo"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton CmdActivar 
      Caption         =   "Activar para exportar"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   6120
      Width           =   2055
   End
   Begin VB.TextBox TxtDesde 
      Height          =   285
      Left            =   3360
      TabIndex        =   6
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox TxtHasta 
      Height          =   285
      Left            =   3360
      TabIndex        =   5
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton CmdExportarSiigoCotrascal 
      Caption         =   "Exportar SIIGO"
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   6120
      Width           =   1815
   End
   Begin VB.TextBox TxtRuta 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   6015
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   9840
      TabIndex        =   0
      Top             =   6600
      Width           =   1335
   End
   Begin MSComctlLib.ListView LstRecibos 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Numero"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tercero"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Rte Fte"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde:"
      Height          =   195
      Left            =   2760
      TabIndex        =   8
      Top             =   6120
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   2760
      TabIndex        =   7
      Top             =   6480
      Width           =   465
   End
   Begin VB.Label LblRuta 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   390
   End
End
Attribute VB_Name = "FrmExportarRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstRecibosExp As New ADODB.Recordset

Private Sub CmdActivar_Click()
  If Val(TxtDesde.Text) <> 0 Then
    If Val(TxtHasta.Text) <> 0 Then
      FufuSt = "UPDATE recibos_caja SET Exportado = 0 WHERE numero >= " & Val(TxtDesde.Text) & " AND numero <= " & Val(TxtHasta.Text)
      AbrirRecorset rstUniversal, FufuSt, CnnPrincipal, adOpenDynamic, adLockOptimistic
      MsgBox "Se han habilidato con exito los recibos", vbInformation
      VerRecibos
    End If
  End If
End Sub

Private Sub CmdExportarContai_Click()
  Dim rstReciboDetalle As New ADODB.Recordset
  rstReciboDetalle.CursorLocation = adUseClient
  Dim RutaSalida As String
  Dim Fila        As Long
  Dim Columna     As Long
  
On Error GoTo Error_Handler
    RutaSalida = TxtRuta.Text & "recexp" & Format(Date, "dd.mm.yy") & "." & Format(Time, "hh.mm.ss") & ".txt"
    Dim J As Integer
    Dim strCuenta As String
    Dim strDetalle As String
    Dim intTipo As Integer
    Dim douValor As Double
    Dim douBase As Double
    Dim strNumero As String
    Dim strNumeroReferencia As String
    Dim intNroRegistros As Integer
    Fila = 0
    II = 1
    Open RutaSalida For Append As #1
    IniProg LstRecibos.ListItems.Count
    Print #1, "Cuenta  Comprobante Fecha(mm/dd/yyyy) Documento Documento Ref.  Nit Detalle Tipo  Valor Base  Centro de Costo Trans. Ext  Plazo"
    While II <= LstRecibos.ListItems.Count
      If LstRecibos.ListItems(II).Checked = True Then
        rstRecibosExp.Open "SELECT recibos_caja.*, terceros.RazonSocial, bancos.cuenta_contable " & _
                            "FROM recibos_caja " & _
                            "LEFT JOIN terceros ON recibos_caja.IdTercero = terceros.IdTercero " & _
                            "LEFT JOIN bancos ON recibos_caja.codigo_banco_fk = bancos.codigo_banco_pk " & _
                            "WHERE Exportado=0 AND IdRecibo = " & LstRecibos.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
                            
        rstReciboDetalle.Open "SELECT recibos_caja_det.*, cuentas_cobrar.NroDocumento, cuentas_cobrar.NroDocumentoOriginal, cuentas_cobrar.TipoFactura, facturas_tipos.Prefijo " & _
                            "FROM recibos_caja_det " & _
                            "LEFT JOIN cuentas_cobrar ON recibos_caja_det.codigo_cuenta_cobrar_fk = cuentas_cobrar.IdCxC " & _
                            "LEFT JOIN facturas_tipos ON cuentas_cobrar.TipoFactura = facturas_tipos.IdTipoFactura " & _
                            "WHERE IdRecibo = " & LstRecibos.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
        While rstReciboDetalle.EOF = False
          strCuenta = "13050501"
          strNumero = Rellenar(rstRecibosExp.Fields("numero"), 9, "0", 1)
          If Val(rstReciboDetalle.Fields("TipoFactura")) = 4 Then
            strNumeroReferencia = Rellenar(rstReciboDetalle.Fields("NroDocumentoOriginal"), 9, "0", 1)
          Else
            strNumeroReferencia = Rellenar(rstReciboDetalle.Fields("Prefijo") & rstReciboDetalle.Fields("NroDocumento"), 9, "0", 1)
          End If
          strDetalle = "PAGO FACTURA"
          intTipo = 2
          douValor = rstReciboDetalle.Fields("valor") + rstReciboDetalle.Fields("retencion_fuente") + rstReciboDetalle.Fields("retencion_ica") + rstReciboDetalle.Fields("descuento")
          Print #1, strCuenta & Chr(9) & "00027" & Chr(9) & Format(rstRecibosExp.Fields("FechaPago"), "mm/dd/yyyy") & Chr(9) & strNumero & Chr(9) & strNumeroReferencia & Chr(9) & rstRecibosExp.Fields("IdTercero") & Chr(9) & strDetalle & Chr(9) & intTipo & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & "0" & Chr(9) & "404" & Chr(9) & "" & Chr(9) & "0"
          
          'Retencion en la fuente
          If Val(rstReciboDetalle.Fields("retencion_fuente")) > 0 Then
            strCuenta = "13551501"
            strNumero = Rellenar(rstRecibosExp.Fields("numero"), 9, "0", 1)
            strDetalle = "RETENCION FUENTE"
            intTipo = 1
            douValor = rstReciboDetalle.Fields("retencion_fuente")
            douBase = rstReciboDetalle.Fields("valor")
            Print #1, strCuenta & Chr(9) & "00027" & Chr(9) & Format(rstRecibosExp.Fields("FechaPago"), "mm/dd/yyyy") & Chr(9) & strNumero & Chr(9) & strNumero & Chr(9) & rstRecibosExp.Fields("IdTercero") & Chr(9) & strDetalle & Chr(9) & intTipo & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & Format(douBase, "##0.00;(##0.00)") & Chr(9) & "404" & Chr(9) & "" & Chr(9) & "0"
          End If
          
          'Retencion Ica
          If Val(rstReciboDetalle.Fields("retencion_ica")) > 0 Then
            strCuenta = "13551801"
            strNumero = Rellenar(rstRecibosExp.Fields("numero"), 9, "0", 1)
            strDetalle = "RETENCION ICA"
            intTipo = 1
            douValor = rstReciboDetalle.Fields("retencion_ica")
            Print #1, strCuenta & Chr(9) & "00027" & Chr(9) & Format(rstRecibosExp.Fields("FechaPago"), "mm/dd/yyyy") & Chr(9) & strNumero & Chr(9) & strNumero & Chr(9) & rstRecibosExp.Fields("IdTercero") & Chr(9) & strDetalle & Chr(9) & intTipo & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & "0" & Chr(9) & "404" & Chr(9) & "" & Chr(9) & "0"
          End If
          
          'Descuento
          If Val(rstReciboDetalle.Fields("descuento")) > 0 Then
            strCuenta = "53053501"
            strNumero = Rellenar(rstRecibosExp.Fields("numero"), 9, "0", 1)
            strDetalle = "DESCUENTO"
            intTipo = 1
            douValor = rstReciboDetalle.Fields("descuento")
            Print #1, strCuenta & Chr(9) & "00027" & Chr(9) & Format(rstRecibosExp.Fields("FechaPago"), "mm/dd/yyyy") & Chr(9) & strNumero & Chr(9) & strNumero & Chr(9) & rstRecibosExp.Fields("IdTercero") & Chr(9) & strDetalle & Chr(9) & intTipo & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & "0" & Chr(9) & "404" & Chr(9) & "" & Chr(9) & "0"
          End If
          rstReciboDetalle.MoveNext
        Wend
        rstReciboDetalle.Close
        
        'Banco
        strCuenta = rstRecibosExp.Fields("cuenta_contable")
        strNumero = Rellenar(rstRecibosExp.Fields("numero"), 9, "0", 1)
        strDetalle = "BANCO"
        intTipo = 1
        douValor = rstRecibosExp.Fields("Total")
        Print #1, strCuenta & Chr(9) & "00027" & Chr(9) & Format(rstRecibosExp.Fields("FechaPago"), "mm/dd/yyyy") & Chr(9) & strNumero & Chr(9) & strNumero & Chr(9) & rstRecibosExp.Fields("IdTercero") & Chr(9) & strDetalle & Chr(9) & intTipo & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & "0" & Chr(9) & "404" & Chr(9) & "" & Chr(9) & "0"
        
        rstRecibosExp.Close
        rstRecibosExp.Open "UPDATE recibos_caja SET Exportado=1 where IdRecibo=" & LstRecibos.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
        LstRecibos.ListItems.Remove (II)
      Else
       II = II + 1
      End If
      Prog (II)
    Wend
    FinProg
    Close #1
  
  Exit Sub
Error_Handler:

        
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical

End Sub

Private Sub CmdExportarSiigoCotrascal_Click()
  Dim rstReciboDetalle As New ADODB.Recordset
  Dim rstCuentaCobrar As New ADODB.Recordset
  rstReciboDetalle.CursorLocation = adUseClient
  rstCuentaCobrar.CursorLocation = adUseClient
  Dim RutaSalida As String
  Dim Fila        As Long
  Dim Columna     As Long
  
  
'On Error GoTo Error_Handler
    RutaSalida = TxtRuta.Text & "recexpsiigo" & Format(Date, "dd.mm.yy") & "." & Format(Time, "hh.mm.ss") & ".txt"
    Dim strSql As String
    Dim intSecuencia As Integer
    Dim strCuenta As String
    Dim strDetalle As String
    Dim strTipo As String
    Dim strComprobante As String
    Dim strNit As String
    Dim strCentroCostos As String
    Dim strVendedor As String
    Dim douValor As Double
    Dim douRetencionFuente As Double
    Dim strValor As String
    Dim strNumero As String
    Dim intNroRegistros As Integer
    Dim strDocumentoCruce As String
    Dim strTipoDocumentoCruce As String
    Fila = 0
    II = 1
    Open RutaSalida For Append As #1
    'Print #1, "Cuenta  Comprobante Fecha(mm/dd/yyyy) Documento Documento Ref.  Nit Detalle Tipo  Valor Base  Centro de Costo Trans. Ext  Plazo"
    While II <= LstRecibos.ListItems.Count
      If LstRecibos.ListItems(II).Checked = True Then
        rstRecibosExp.Open "SELECT recibos_caja.*, terceros.RazonSocial, bancos.cuenta_contable as cuentaBanco, bancos.nombre as nombreBanco " & _
                            "FROM recibos_caja " & _
                            "LEFT JOIN terceros ON recibos_caja.IdTercero = terceros.IdTercero " & _
                            "LEFT JOIN bancos ON recibos_caja.codigo_banco_fk = bancos.codigo_banco_pk " & _
                            "WHERE Exportado=0 AND IdRecibo = " & LstRecibos.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
        If rstRecibosExp.RecordCount > 0 Then
          strComprobante = "001"
          strNumero = Rellenar(rstRecibosExp.Fields("numero"), 11, "0", 1)
          intSecuencia = 1
          strNit = rstRecibosExp!IdTercero
          strCentroCostos = "0001"
          strVendedor = "0001"
          strSql = "Select recibos_caja_det.*, cuentas_cobrar.NroDocumento, cuentas_cobrar.FhVence, cuentas_cobrar.TipoFactura from recibos_caja_det left join cuentas_cobrar ON recibos_caja_det.codigo_cuenta_cobrar_fk = cuentas_cobrar.IdCxC where IdRecibo = " & rstRecibosExp!IdRecibo
          AbrirRecorset rstReciboDetalle, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
          Do While rstReciboDetalle.EOF = False
          
            strSql = "SELECT centrosoperaciones.cuenta_cartera, cuentas_cobrar.TipoFactura " & _
                    "FROM cuentas_cobrar " & _
                    "LEFT JOIN centrosoperaciones ON cuentas_cobrar.IdPO = centrosoperaciones.IDPO " & _
                    "WHERE IdCxC=" & rstReciboDetalle.Fields("codigo_cuenta_cobrar_fk")
            
            AbrirRecorset rstCuentaCobrar, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
              If Val(rstCuentaCobrar.Fields("TipoFactura")) = 1 Then
                strCuenta = rstCuentaCobrar.Fields("cuenta_cartera")
              End If
              If Val(rstCuentaCobrar.Fields("TipoFactura")) = 2 Then
                strCuenta = "13050501"
              End If
              If Val(rstCuentaCobrar.Fields("TipoFactura")) = 3 Then
                strCuenta = "13050502"
              End If
              
            CerrarRecorset rstCuentaCobrar
            'Cuenta cliente
            strDetalle = "CANC FACT " & rstReciboDetalle!NroDocumento
            strTipo = "C"
            douValor = rstReciboDetalle.Fields("valor") + rstReciboDetalle.Fields("retencion_ica") + rstReciboDetalle.Fields("retencion_fuente") + rstReciboDetalle.Fields("descuento") + rstReciboDetalle.Fields("ajuste_peso")
            douValor = Round(douValor)
            strValor = Limpiar(Format(douValor, "##0.00;(##0.00)") & "")
            'Corriente
            If Val(rstReciboDetalle!TipoFactura) = 1 Or Val(rstReciboDetalle!TipoFactura) = 4 Then
              strTipoDocumentoCruce = "F003"
            End If
            'Contado
            If Val(rstReciboDetalle!TipoFactura) = 2 Then
              strTipoDocumentoCruce = "F001"
            End If
            'Destino
            If Val(rstReciboDetalle!TipoFactura) = 3 Then
              strTipoDocumentoCruce = "F002"
            End If
            
            strDocumentoCruce = strTipoDocumentoCruce & Rellenar(rstReciboDetalle!NroDocumento, 11, "0", 1) & Rellenar(intSecuencia & "", 3, "0", 1) & Format(rstReciboDetalle!FhVence, "yyyymmdd") & "0001" & "00"
            Print #1, "R" & strComprobante & strNumero & Rellenar(intSecuencia & "", 5, "0", 1) & Rellenar(strNit, 13, "0", 1) & "000" & strCuenta & "000000000000000" & Format(rstRecibosExp!FechaPago, "yyyymmdd") & strCentroCostos & "000" & Rellenar(strDetalle, 50, " ", 0) & strTipo & Rellenar(strValor, 15, "0", 1) & "000000000000000" & strVendedor & "0001" & "001" & "0001" & "000" & "000000000000000" & strDocumentoCruce
            intSecuencia = intSecuencia + 1
            
            'Cuenta Rte Ica
            If Val(rstReciboDetalle.Fields("retencion_ica")) > 0 Then
              strCuenta = "1355180100"
              strDetalle = "RTE ICA  " & rstReciboDetalle!NroDocumento
              strTipo = "D"
              douValor = rstReciboDetalle.Fields("retencion_ica")
              douValor = Round(douValor)
              strValor = Limpiar(Format(douValor, "##0.00;(##0.00)") & "")
              Print #1, "R" & strComprobante & strNumero & Rellenar(intSecuencia & "", 5, "0", 1) & Rellenar(strNit, 13, "0", 1) & "000" & strCuenta & "0000000000000" & Format(rstRecibosExp!FechaPago, "yyyymmdd") & strCentroCostos & "000" & Rellenar(strDetalle, 50, " ", 0) & strTipo & Rellenar(strValor, 15, "0", 1) & "000000000000000" & strVendedor & "0001" & "001" & "0001" & "000" & "000000000000000" & strDocumentoCruce
              intSecuencia = intSecuencia + 1
            End If
            
            'Cuenta Rte fte
            If Val(rstReciboDetalle.Fields("retencion_fuente")) > 0 Then
              strCuenta = "1355150200"
              strDetalle = "RTE FUENTE " & rstReciboDetalle!NroDocumento
              strTipo = "D"
              douValor = rstReciboDetalle.Fields("retencion_fuente")
              douValor = Round(douValor)
              strValor = Limpiar(Format(douValor, "##0.00;(##0.00)") & "")
              Print #1, "R" & strComprobante & strNumero & Rellenar(intSecuencia & "", 5, "0", 1) & Rellenar(strNit, 13, "0", 1) & "000" & strCuenta & "0000000000000" & Format(rstRecibosExp!FechaPago, "yyyymmdd") & strCentroCostos & "000" & Rellenar(strDetalle, 50, " ", 0) & strTipo & Rellenar(strValor, 15, "0", 1) & "000000000000000" & strVendedor & "0001" & "001" & "0001" & "000" & "000000000000000" & strDocumentoCruce
              intSecuencia = intSecuencia + 1
            End If
            
            'Descuento
            If Val(rstReciboDetalle.Fields("descuento")) > 0 Then
              strCuenta = "5305350100"
              strDetalle = "DESCUENTO " & rstReciboDetalle!NroDocumento
              strTipo = "D"
              douValor = rstReciboDetalle.Fields("descuento")
              douValor = Round(douValor)
              strValor = Limpiar(Format(douValor, "##0.00;(##0.00)") & "")
              Print #1, "R" & strComprobante & strNumero & Rellenar(intSecuencia & "", 5, "0", 1) & Rellenar(strNit, 13, "0", 1) & "000" & strCuenta & "0000000000000" & Format(rstRecibosExp!FechaPago, "yyyymmdd") & strCentroCostos & "000" & Rellenar(strDetalle, 50, " ", 0) & strTipo & Rellenar(strValor, 15, "0", 1) & "000000000000000" & strVendedor & "0001" & "001" & "0001" & "000" & "000000000000000" & strDocumentoCruce
              intSecuencia = intSecuencia + 1
            End If
            
            'Ajuste al peso
            If Val(rstReciboDetalle.Fields("ajuste_peso")) <> 0 Then
              If Val(rstReciboDetalle.Fields("ajuste_peso")) > 0 Then
                strCuenta = "5395950100"
                strTipo = "D"
              Else
                strCuenta = "4295810100"
                strTipo = "C"
              End If
              strDetalle = "AJUSTE PESO " & rstReciboDetalle!NroDocumento
              douValor = rstReciboDetalle.Fields("ajuste_peso")
              douValor = Round(douValor)
              strValor = Limpiar(Format(douValor, "##0.00;##0.00") & "")
              Print #1, "R" & strComprobante & strNumero & Rellenar(intSecuencia & "", 5, "0", 1) & Rellenar(strNit, 13, "0", 1) & "000" & strCuenta & "0000000000000" & Format(rstRecibosExp!FechaPago, "yyyymmdd") & strCentroCostos & "000" & Rellenar(strDetalle, 50, " ", 0) & strTipo & Rellenar(strValor, 15, "0", 1) & "000000000000000" & strVendedor & "0001" & "001" & "0001" & "000" & "000000000000000" & strDocumentoCruce
              intSecuencia = intSecuencia + 1
            End If
            
            rstReciboDetalle.MoveNext
          Loop
          'Banco
          strCuenta = rstRecibosExp.Fields("cuentaBanco")
          strDetalle = rstRecibosExp.Fields("nombreBanco")
          strTipo = "D"
          douValor = rstRecibosExp.Fields("Total")
          douValor = Round(douValor)
          strValor = Limpiar(Format(douValor, "##0.00;(##0.00)") & "")
          strDocumentoCruce = "R001" & Rellenar(rstRecibosExp!Numero, 11, "0", 1) & Rellenar(intSecuencia & "", 3, "0", 1) & Format(rstRecibosExp!FechaPago, "yyyymmdd") & "0001" & "00"
          Print #1, "R" & strComprobante & strNumero & Rellenar(intSecuencia & "", 5, "0", 1) & Rellenar(strNit, 13, "0", 1) & "000" & strCuenta & "0000000000000" & Format(rstRecibosExp!FechaPago, "yyyymmdd") & strCentroCostos & "000" & Rellenar(strDetalle, 50, " ", 0) & strTipo & Rellenar(strValor, 15, "0", 1) & "000000000000000" & strVendedor & "0001" & "001" & "0001" & "000" & "000000000000000" & strDocumentoCruce
          CerrarRecorset rstReciboDetalle
        End If
     
        rstRecibosExp.Close
        rstRecibosExp.Open "UPDATE recibos_caja SET Exportado=1 where IdRecibo=" & LstRecibos.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
        LstRecibos.ListItems.Remove (II)
      Else
       II = II + 1
      End If
    Wend
    Close #1
  
  Exit Sub
'Error_Handler:
    'If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub CmdSeleccionarTodo_Click()
  II = 1
  For II = 1 To LstRecibos.ListItems.Count
    LstRecibos.ListItems(II).Checked = True
  Next
End Sub

Private Sub Form_Load()
  rstRecibosExp.CursorLocation = adUseClient
  TxtRuta.Text = GetSetting("Kit Logistics", "Facturacion", "RutaExportarArchivoFacturas")
  VerRecibos
End Sub

Private Sub VerRecibos()
  Dim strSql As String
  LstRecibos.ListItems.Clear
  strSql = "SELECT recibos_caja.*, terceros.RazonSocial " & _
                          "FROM recibos_caja " & _
                          "LEFT JOIN terceros ON recibos_caja.IdTercero = terceros.IdTercero " & _
                          "WHERE Exportado=0 AND Impreso = 1 AND numero <> 0 order by FechaPago"
  rstRecibosExp.Open strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
  IniProg rstRecibosExp.RecordCount
  If rstRecibosExp.RecordCount > 0 Then
    Do While rstRecibosExp.EOF = False
      Prog (rstRecibosExp.AbsolutePosition)
      Set Item = LstRecibos.ListItems.Add(, , rstRecibosExp!IdRecibo)
      Item.SubItems(1) = rstRecibosExp.Fields("numero")
      Item.SubItems(2) = Format(rstRecibosExp!FechaPago, "dd/mm/yy")
      Item.SubItems(3) = rstRecibosExp!RazonSocial & ""
      Item.SubItems(4) = rstRecibosExp!total & ""
      Item.SubItems(5) = rstRecibosExp.Fields("retencion_fuente") & ""
      rstRecibosExp.MoveNext
    Loop
  End If
  FinProg
  rstRecibosExp.Close
End Sub

Private Function Rellenar(Dato As String, Tama�o As Integer, Caracter As String, Posicion As Byte) As String
  FufuSt = ""
  If Len(Dato) <= Tama�o Then
    For FufuLo = 1 To Tama�o - Len(Dato)
      FufuSt = FufuSt & Caracter
    Next
    If Posicion = 1 Then
      Rellenar = FufuSt & Dato
    Else
      Rellenar = Dato & FufuSt
    End If
  End If
End Function

Private Function Limpiar(Dato As String) As String
  FufuSt = ""
  If Len(Dato) > 0 Then
    For FufuLo = 1 To Len(Dato)
      If Mid(Dato, FufuLo, 1) <> "." Then
        FufuSt = FufuSt & Mid(Dato, FufuLo, 1)
      End If
    Next
  End If
  Limpiar = FufuSt
End Function
