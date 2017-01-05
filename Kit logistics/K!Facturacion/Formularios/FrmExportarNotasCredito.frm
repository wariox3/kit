VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmExportarNotasCredito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar notas credito"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   13035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   11520
      TabIndex        =   6
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox TxtRuta 
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   6015
   End
   Begin VB.CommandButton CmdExportarSiigoCotrascal 
      Caption         =   "Exportar SIIGO"
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   6120
      Width           =   1815
   End
   Begin VB.TextBox TxtHasta 
      Height          =   285
      Left            =   3360
      TabIndex        =   3
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox TxtDesde 
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton CmdActivar 
      Caption         =   "Activar para exportar"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton CmdSeleccionarTodo 
      Caption         =   "Seleccionar todo"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   1455
   End
   Begin MSComctlLib.ListView LstNotasCredito 
      Height          =   5535
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   12735
      _ExtentX        =   22463
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
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tercero"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Concepto"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label LblRuta 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   2760
      TabIndex        =   9
      Top             =   6480
      Width           =   465
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
End
Attribute VB_Name = "FrmExportarNotasCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstNotasCredito As New ADODB.Recordset

Private Sub CmdActivar_Click()
  If Val(TxtDesde.Text) <> 0 Then
    If Val(TxtHasta.Text) <> 0 Then
      FufuSt = "UPDATE notas_credito SET Exportada = 0 WHERE numeroNotaCredito >= " & Val(TxtDesde.Text) & " AND numeroNotaCredito <= " & Val(TxtHasta.Text)
      AbrirRecorset rstUniversal, FufuSt, CnnPrincipal, adOpenDynamic, adLockOptimistic
      MsgBox "Se han habilidato con exito las notas credito", vbInformation
      VerRecibos
    End If
  End If
End Sub

Private Sub CmdExportarSiigoCotrascal_Click()
  Dim rstNotaCreditoDetalle As New ADODB.Recordset
  Dim rstCuentaCobrar As New ADODB.Recordset
  
  rstNotaCreditoDetalle.CursorLocation = adUseClient
  rstCuentaCobrar.CursorLocation = adUseClient
  
  Dim RutaSalida As String
  Dim Fila        As Long
  Dim Columna     As Long
  
  
'On Error GoTo Error_Handler
    RutaSalida = TxtRuta.Text & "notacreditoexpsiigo" & Format(Date, "dd.mm.yy") & "." & Format(Time, "hh.mm.ss") & ".txt"
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
    Fila = 0
    II = 1
    Open RutaSalida For Append As #1
    'Print #1, "Cuenta  Comprobante Fecha(mm/dd/yyyy) Documento Documento Ref.  Nit Detalle Tipo  Valor Base  Centro de Costo Trans. Ext  Plazo"
    While II <= LstNotasCredito.ListItems.Count
      If LstNotasCredito.ListItems(II).Checked = True Then
        rstNotasCredito.Open "SELECT notas_credito.*, terceros.RazonSocial, nota_credito_tipo.Cuenta as cuentaConcepto, nota_credito_tipo.Nombre as nombreConcepto, nota_credito_tipo.anulacion as anulacion " & _
                            "FROM notas_credito " & _
                            "LEFT JOIN terceros ON notas_credito.IdTercero = terceros.IdTercero " & _
                            "LEFT JOIN nota_credito_tipo ON notas_credito.IdNotaCreditoTipo = nota_credito_tipo.IdNotaCreditoTipo " & _
                            "WHERE Exportada=0 AND IdNotaCredito = " & LstNotasCredito.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
        If rstNotasCredito.RecordCount > 0 Then
          strComprobante = "001"
          strNumero = Rellenar(rstNotasCredito.Fields("numeroNotaCredito"), 11, "0", 1)
          intSecuencia = 1
          strNit = rstNotasCredito!IdTercero
          strCentroCostos = "0001"
          strVendedor = "0001"
          If Val(rstNotasCredito.Fields("anulacion")) = 0 Then
            strSql = "Select notas_credito_det.*, cuentas_cobrar.NroDocumento, cuentas_cobrar.FhVence from notas_credito_det left join cuentas_cobrar ON notas_credito_det.IdCxC = cuentas_cobrar.IdCxC where IdNotaCredito = " & rstNotasCredito!IdNotaCredito
            AbrirRecorset rstNotaCreditoDetalle, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
            Do While rstNotaCreditoDetalle.EOF = False
              'Cuenta cliente
              strCuenta = "1305050300"
              strDetalle = "CANC FACT " & rstNotaCreditoDetalle!NroDocumento
              strTipo = "C"
              douValor = rstNotaCreditoDetalle.Fields("valor")
              douValor = Round(douValor)
              strValor = Limpiar(Format(douValor, "##0.00;(##0.00)") & "")
              strDocumentoCruce = "F003" & Rellenar(rstNotaCreditoDetalle!NroDocumento, 11, "0", 1) & Rellenar(intSecuencia & "", 3, "0", 1) & Format(rstNotaCreditoDetalle!FhVence, "yyyymmdd") & "0001" & "00"
              Print #1, "C" & strComprobante & strNumero & Rellenar(intSecuencia & "", 5, "0", 1) & Rellenar(strNit, 13, "0", 1) & "000" & strCuenta & "0000000000000" & Format(rstNotasCredito!Fecha, "yyyymmdd") & strCentroCostos & "000" & Rellenar(strDetalle, 50, " ", 0) & strTipo & Rellenar(strValor, 15, "0", 1) & "000000000000000" & strVendedor & "0001" & "001" & "0001" & "000" & "000000000000000" & strDocumentoCruce
              intSecuencia = intSecuencia + 1
              rstNotaCreditoDetalle.MoveNext
            Loop
            'Cruce
            strCuenta = rstNotasCredito.Fields("cuentaConcepto")
            strDetalle = rstNotasCredito.Fields("nombreConcepto")
            strTipo = "D"
            douValor = rstNotasCredito.Fields("Total")
            douValor = Round(douValor)
            strValor = Limpiar(Format(douValor, "##0.00;(##0.00)") & "")
            strDocumentoCruce = "R001" & Rellenar(rstNotasCredito!numeroNotaCredito, 11, "0", 1) & Rellenar(intSecuencia & "", 3, "0", 1) & Format(rstNotasCredito!Fecha, "yyyymmdd") & "0001" & "00"
            Print #1, "C" & strComprobante & strNumero & Rellenar(intSecuencia & "", 5, "0", 1) & Rellenar(strNit, 13, "0", 1) & "000" & strCuenta & "0000000000000" & Format(rstNotasCredito!Fecha, "yyyymmdd") & strCentroCostos & "000" & Rellenar(strDetalle, 50, " ", 0) & strTipo & Rellenar(strValor, 15, "0", 1) & "000000000000000" & strVendedor & "0001" & "001" & "0001" & "000" & "000000000000000" & strDocumentoCruce
            CerrarRecorset rstNotaCreditoDetalle
          End If
          
          If Val(rstNotasCredito.Fields("anulacion")) = 1 Then
            strSql = "Select notas_credito_det.*, cuentas_cobrar.NroDocumento, cuentas_cobrar.FhVence from notas_credito_det left join cuentas_cobrar ON notas_credito_det.IdCxC = cuentas_cobrar.IdCxC where IdNotaCredito = " & rstNotasCredito!IdNotaCredito
            AbrirRecorset rstNotaCreditoDetalle, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
            Do While rstNotaCreditoDetalle.EOF = False
              strSql = "Select cuentas_cobrar.* from cuentas_cobrar Where IdCxC = " & rstNotaCreditoDetalle.Fields("IdCxC")
              AbrirRecorset rstCuentaCobrar, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
              
              strCuenta = "4145050100"
              strDetalle = "ANULACION FLETE " & rstNotaCreditoDetalle!NroDocumento
              strTipo = "D"
              douValor = rstCuentaCobrar.Fields("VrFlete")
              douValor = Round(douValor)
              strValor = Limpiar(Format(douValor, "##0.00;(##0.00)") & "")
              strDocumentoCruce = "F003" & Rellenar(rstNotaCreditoDetalle!NroDocumento, 11, "0", 1) & Rellenar(intSecuencia & "", 3, "0", 1) & Format(rstNotaCreditoDetalle!FhVence, "yyyymmdd") & "0001" & "00"
              Print #1, "C" & strComprobante & strNumero & Rellenar(intSecuencia & "", 5, "0", 1) & Rellenar(strNit, 13, "0", 1) & "000" & strCuenta & "0000000000000" & Format(rstNotasCredito!Fecha, "yyyymmdd") & strCentroCostos & "000" & Rellenar(strDetalle, 50, " ", 0) & strTipo & Rellenar(strValor, 15, "0", 1) & "000000000000000" & strVendedor & "0001" & "001" & "0001" & "000" & "000000000000000" & strDocumentoCruce
              intSecuencia = intSecuencia + 1
              
              strCuenta = "4145950100"
              strDetalle = "ANULACION MANEJO " & rstNotaCreditoDetalle!NroDocumento
              strTipo = "D"
              douValor = rstCuentaCobrar.Fields("VrManejo")
              douValor = Round(douValor)
              strValor = Limpiar(Format(douValor, "##0.00;(##0.00)") & "")
              strDocumentoCruce = "F003" & Rellenar(rstNotaCreditoDetalle!NroDocumento, 11, "0", 1) & Rellenar(intSecuencia & "", 3, "0", 1) & Format(rstNotaCreditoDetalle!FhVence, "yyyymmdd") & "0001" & "00"
              Print #1, "C" & strComprobante & strNumero & Rellenar(intSecuencia & "", 5, "0", 1) & Rellenar(strNit, 13, "0", 1) & "000" & strCuenta & "0000000000000" & Format(rstNotasCredito!Fecha, "yyyymmdd") & strCentroCostos & "000" & Rellenar(strDetalle, 50, " ", 0) & strTipo & Rellenar(strValor, 15, "0", 1) & "000000000000000" & strVendedor & "0001" & "001" & "0001" & "000" & "000000000000000" & strDocumentoCruce
              intSecuencia = intSecuencia + 1
              
              'Cuenta cliente
              strCuenta = "1305050300"
              strDetalle = "CANC FACT " & rstNotaCreditoDetalle!NroDocumento
              strTipo = "C"
              douValor = rstNotaCreditoDetalle.Fields("valor")
              douValor = Round(douValor)
              strValor = Limpiar(Format(douValor, "##0.00;(##0.00)") & "")
              strDocumentoCruce = "F003" & Rellenar(rstNotaCreditoDetalle!NroDocumento, 11, "0", 1) & Rellenar(intSecuencia & "", 3, "0", 1) & Format(rstNotaCreditoDetalle!FhVence, "yyyymmdd") & "0001" & "00"
              Print #1, "C" & strComprobante & strNumero & Rellenar(intSecuencia & "", 5, "0", 1) & Rellenar(strNit, 13, "0", 1) & "000" & strCuenta & "0000000000000" & Format(rstNotasCredito!Fecha, "yyyymmdd") & strCentroCostos & "000" & Rellenar(strDetalle, 50, " ", 0) & strTipo & Rellenar(strValor, 15, "0", 1) & "000000000000000" & strVendedor & "0001" & "001" & "0001" & "000" & "000000000000000" & strDocumentoCruce
              intSecuencia = intSecuencia + 1
              rstNotaCreditoDetalle.MoveNext
            Loop
            'Cruce
            'strCuenta = rstNotasCredito.Fields("cuentaConcepto")
            'strDetalle = rstNotasCredito.Fields("nombreConcepto")
            'strTipo = "D"
            'douValor = rstNotasCredito.Fields("Total")
            'douValor = Round(douValor)
            'strValor = Limpiar(Format(douValor, "##0.00;(##0.00)") & "")
            'strDocumentoCruce = "R001" & Rellenar(rstNotasCredito!numeroNotaCredito, 11, "0", 1) & Rellenar(intSecuencia & "", 3, "0", 1) & Format(rstNotasCredito!Fecha, "yyyymmdd") & "0001" & "00"
            'Print #1, "C" & strComprobante & strNumero & Rellenar(intSecuencia & "", 5, "0", 1) & Rellenar(strNit, 13, "0", 1) & "000" & strCuenta & "0000000000000" & Format(rstNotasCredito!Fecha, "yyyymmdd") & strCentroCostos & "000" & Rellenar(strDetalle, 50, " ", 0) & strTipo & Rellenar(strValor, 15, "0", 1) & "000000000000000" & strVendedor & "0001" & "001" & "0001" & "000" & "000000000000000" & strDocumentoCruce
            CerrarRecorset rstNotaCreditoDetalle
          End If
          
        End If
     
        rstNotasCredito.Close
        rstNotasCredito.Open "UPDATE notas_credito SET Exportada=1 where IdNotaCredito=" & LstNotasCredito.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
        LstNotasCredito.ListItems.Remove (II)
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
  For II = 1 To LstNotasCredito.ListItems.Count
    LstNotasCredito.ListItems(II).Checked = True
  Next
End Sub

Private Sub Form_Load()
  rstNotasCredito.CursorLocation = adUseClient
  TxtRuta.Text = GetSetting("Kit Logistics", "Facturacion", "RutaExportarArchivoFacturas")
  VerRecibos
End Sub

Private Sub VerRecibos()
  Dim strSql As String
  LstNotasCredito.ListItems.Clear
  strSql = "SELECT notas_credito.*, terceros.RazonSocial, nota_credito_tipo.Nombre as Tipo " & _
                          "FROM notas_credito " & _
                          "LEFT JOIN terceros ON notas_credito.IdTercero = terceros.IdTercero " & _
                          "LEFT JOIN nota_credito_tipo ON notas_credito.IdNotaCreditoTipo = nota_credito_tipo.IdNotaCreditoTipo " & _
                          "WHERE Exportada=0 AND Impreso = 1 AND numeroNotaCredito <> 0"
  rstNotasCredito.Open strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
  IniProg rstNotasCredito.RecordCount
  If rstNotasCredito.RecordCount > 0 Then
    Do While rstNotasCredito.EOF = False
      Prog (rstNotasCredito.AbsolutePosition)
      Set Item = LstNotasCredito.ListItems.Add(, , rstNotasCredito!IdNotaCredito)
      Item.SubItems(1) = rstNotasCredito.Fields("numeroNotaCredito")
      Item.SubItems(2) = Format(rstNotasCredito!Fecha, "dd/mm/yy")
      Item.SubItems(3) = rstNotasCredito!RazonSocial & ""
      Item.SubItems(4) = rstNotasCredito!Tipo & ""
      Item.SubItems(5) = rstNotasCredito!Total & ""
      rstNotasCredito.MoveNext
    Loop
  End If
  FinProg
  rstNotasCredito.Close
End Sub

Private Function Rellenar(Dato As String, Tamaño As Integer, Caracter As String, Posicion As Byte) As String
  FufuSt = ""
  If Len(Dato) < Tamaño Then
    For FufuLo = 1 To Tamaño - Len(Dato)
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

