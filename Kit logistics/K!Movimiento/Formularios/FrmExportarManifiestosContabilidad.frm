VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmExportarManifiestosContabilidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar manifiestos contabilidad"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   12825
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtRuta 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Text            =   "C:\ExportarManifiestos\"
      Top             =   240
      Width           =   8055
   End
   Begin VB.CommandButton CmdExportar 
      Caption         =   "Exportar"
      Height          =   255
      Left            =   10560
      TabIndex        =   1
      Top             =   5520
      Width           =   2055
   End
   Begin MSComctlLib.ListView LstDespachos 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   8493
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Orden"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Man"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Placa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "P"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Flete"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "TotalCE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "TotalRec"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   390
   End
End
Attribute VB_Name = "FrmExportarManifiestosContabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstDespachosExp As New ADODB.Recordset
Private Sub VerDespachos()
  LstDespachos.ListItems.Clear
  rstDespachosExp.Open "Select despachos.*, vehiculos.VehiculoPropio from despachos LEFT JOIN  vehiculos ON despachos.IdVehiculo = vehiculos.IdPlaca where ExportadoContabilidad=0 and (Estado='I' or Estado='G' or Estado='V') and IdManifiesto<>0 order by IdManifiesto", CnnPrincipal, adOpenDynamic, adLockOptimistic
  IniProg 1, rstDespachosExp.RecordCount
  If rstDespachosExp.RecordCount > 0 Then
    Do While rstDespachosExp.EOF = False
      Prog (rstDespachosExp.AbsolutePosition)
      Set Item = LstDespachos.ListItems.Add(, , rstDespachosExp!OrdDespacho)
      Item.SubItems(1) = rstDespachosExp!IdManifiesto
      Item.SubItems(2) = Format(rstDespachosExp!FhExpedicion, "dd/mm/yy")
      Item.SubItems(3) = rstDespachosExp!IdVehiculo
      If Val(rstDespachosExp!VehiculoPropio) = 1 Then
        Item.SubItems(4) = "S"
      Else
        Item.SubItems(4) = "N"
      End If
      
      Item.SubItems(6) = rstDespachosExp!VrFlete
      Item.SubItems(7) = rstDespachosExp!TotalCE
      Item.SubItems(7) = rstDespachosExp!TRecaudo
      rstDespachosExp.MoveNext
    Loop
  End If
  FinProg
  rstDespachosExp.Close
End Sub

Private Sub CmdExportar_Click()
  Dim RutaSalida As String
  Dim douValor As Double
  Dim douBase As Double
  Dim strNumero As String
  Dim strNit As String
  Dim strCuenta As String
  Dim strSql As String
'On Error GoTo Error_Handler
    
    RutaSalida = TxtRuta.Text & "manexp" & Format(Date, "dd.mm.yy") & "." & Format(Time, "hh.mm.ss") & ".txt"
    II = 1
    Open RutaSalida For Append As #1
    
    Print #1, "Cuenta  Comprobante Fecha(mm/dd/yyyy) Documento Documento Ref.  Nit Detalle Tipo  Valor Base  Centro de Costo Trans. Ext  Plazo"
    
    While II <= LstDespachos.ListItems.Count
      If LstDespachos.ListItems(II).Checked = True Then
        
        rstDespachosExp.Open "SELECT despachos.*, vehiculos.IdPropietario, vehiculos.VehiculoPropio " & _
                            "FROM despachos " & _
                            "LEFT JOIN  vehiculos ON despachos.IdVehiculo = vehiculos.IdPlaca " & _
                            "WHERE OrdDespacho = " & LstDespachos.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
        strNumero = Rellenar(rstDespachosExp.Fields("IdManifiesto"), 9, "0", 1)
        strNit = rstDespachosExp.Fields("IdPropietario")
        
        'Flete

        If Val(rstDespachosExp.Fields("VehiculoPropio")) = 0 Then
          douValor = rstDespachosExp.Fields("VrFlete")
          douBase = rstDespachosExp.Fields("VrFlete")
        Else
          douValor = 0
        End If
        Print #1, "41450520" & Chr(9) & "00025" & Chr(9) & Format(rstDespachosExp.Fields("FhExpedicion"), "mm/dd/yyyy") & Chr(9) & strNumero & Chr(9) & strNumero & Chr(9) & strNit & Chr(9) & "Flete" & Chr(9) & "1" & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & "0" & Chr(9) & "111" & Chr(9) & "" & Chr(9) & "0"
        If Val(rstDespachosExp.Fields("VehiculoPropio")) = 0 Then
          'Retencion en la fuente
          douValor = rstDespachosExp.Fields("VrDctoRteFte")
          Print #1, "13551501" & Chr(9) & "00025" & Chr(9) & Format(rstDespachosExp.Fields("FhExpedicion"), "mm/dd/yyyy") & Chr(9) & strNumero & Chr(9) & strNumero & Chr(9) & strNit & Chr(9) & "Rte Fuente" & Chr(9) & "2" & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & Format(douBase, "##0.00;(##0.00)") & Chr(9) & "" & Chr(9) & "" & Chr(9) & "0"
        End If
        If Val(rstDespachosExp.Fields("VehiculoPropio")) = 0 Then
          'ica
          douValor = rstDespachosExp.Fields("VrDctoIndCom")
          Print #1, "13551802" & Chr(9) & "00025" & Chr(9) & Format(rstDespachosExp.Fields("FhExpedicion"), "mm/dd/yyyy") & Chr(9) & strNumero & Chr(9) & strNumero & Chr(9) & strNit & Chr(9) & "ica" & Chr(9) & "2" & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & Format(douBase, "##0.00;(##0.00)") & Chr(9) & "" & Chr(9) & "" & Chr(9) & "0"
        End If
        If Val(rstDespachosExp.Fields("VehiculoPropio")) = 0 Then
          'Acompa�amiento
          douValor = rstDespachosExp.Fields("VrDctoSeguridad")
          Print #1, "42505015" & Chr(9) & "00025" & Chr(9) & Format(rstDespachosExp.Fields("FhExpedicion"), "mm/dd/yyyy") & Chr(9) & strNumero & Chr(9) & strNumero & Chr(9) & strNit & Chr(9) & "Acompa�amiento" & Chr(9) & "2" & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & "0" & Chr(9) & "111" & Chr(9) & "" & Chr(9) & "0"
        End If
        If Val(rstDespachosExp.Fields("VehiculoPropio")) = 0 Then
          'Cargue
          douValor = rstDespachosExp.Fields("VrDctoCargue")
          Print #1, "42505020" & Chr(9) & "00025" & Chr(9) & Format(rstDespachosExp.Fields("FhExpedicion"), "mm/dd/yyyy") & Chr(9) & strNumero & Chr(9) & strNumero & Chr(9) & strNit & Chr(9) & "Cargue" & Chr(9) & "2" & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & "0" & Chr(9) & "111" & Chr(9) & "" & Chr(9) & "0"
        End If
        If Val(rstDespachosExp.Fields("VehiculoPropio")) = 0 Then
          'Estampilla
          douValor = rstDespachosExp.Fields("VrDctoEstampilla")
          Print #1, "42505025" & Chr(9) & "00025" & Chr(9) & Format(rstDespachosExp.Fields("FhExpedicion"), "mm/dd/yyyy") & Chr(9) & strNumero & Chr(9) & strNumero & Chr(9) & strNit & Chr(9) & "Estampilla" & Chr(9) & "2" & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & "0" & Chr(9) & "111" & Chr(9) & "" & Chr(9) & "0"
        End If
        If Val(rstDespachosExp.Fields("VehiculoPropio")) = 0 Then
          'Papeleria
          douValor = rstDespachosExp.Fields("VrDctoPapeleria")
          Print #1, "42505010" & Chr(9) & "00025" & Chr(9) & Format(rstDespachosExp.Fields("FhExpedicion"), "mm/dd/yyyy") & Chr(9) & strNumero & Chr(9) & strNumero & Chr(9) & strNit & Chr(9) & "Papeleria" & Chr(9) & "2" & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & "0" & Chr(9) & "111" & Chr(9) & "" & Chr(9) & "0"
        End If
        If Val(rstDespachosExp.Fields("VehiculoPropio")) = 0 Then
          'Anticipo
          douValor = rstDespachosExp.Fields("VrAnticipo")
          strCuenta = "13309502"
          If Val(rstDespachosExp.Fields("VehiculoPropio")) = 0 Then
            strCuenta = "13301001"
          End If
          Print #1, strCuenta & Chr(9) & "00025" & Chr(9) & Format(rstDespachosExp.Fields("FhExpedicion"), "mm/dd/yyyy") & Chr(9) & strNumero & Chr(9) & strNumero & Chr(9) & strNit & Chr(9) & "Anticipo" & Chr(9) & "2" & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & "0" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "0"
        End If
        
        strSql = "Select*from sql_im_contraentregas where IdDespacho = " & LstDespachos.ListItems(II)
        AbrirRecorset rstUniversal, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
        Dim strNumeroDoc As String
        Do While rstUniversal.EOF = False
          If Val(rstUniversal.Fields("TipoCobro")) = 2 Then
            'Destino
            douValor = (rstUniversal.Fields("VrFlete") + rstUniversal.Fields("VrManejo")) - Val(rstUniversal.Fields("Abonos") & "")
            strNumeroDoc = Rellenar("A" & rstUniversal.Fields("Guia"), 9, "0", 1)
            Print #1, "13050502" & Chr(9) & "00025" & Chr(9) & Format(rstDespachosExp.Fields("FhExpedicion"), "mm/dd/yyyy") & Chr(9) & strNumero & Chr(9) & strNumeroDoc & Chr(9) & rstUniversal.Fields("Cuenta") & Chr(9) & "Destino" & Chr(9) & "2" & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & "0" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "0"
          End If
          If Val(rstUniversal.Fields("Recaudo")) > 0 Then
            'Recaudo
            douValor = rstUniversal.Fields("Recaudo")
            strNumeroDoc = Rellenar(rstUniversal.Fields("Guia"), 9, "0", 1)
            Print #1, "13050503" & Chr(9) & "00025" & Chr(9) & Format(rstDespachosExp.Fields("FhExpedicion"), "mm/dd/yyyy") & Chr(9) & strNumero & Chr(9) & strNumeroDoc & Chr(9) & rstUniversal.Fields("Cuenta") & Chr(9) & "Recaudo" & Chr(9) & "2" & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & "0" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "0"
            Print #1, "13050503" & Chr(9) & "00024" & Chr(9) & Format(rstUniversal.Fields("FhEntradaBodega"), "mm/dd/yyyy") & Chr(9) & strNumeroDoc & Chr(9) & strNumeroDoc & Chr(9) & rstUniversal.Fields("Cuenta") & Chr(9) & "Recaudo" & Chr(9) & "1" & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & "0" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "0"
            Print #1, "28150550" & Chr(9) & "00024" & Chr(9) & Format(rstUniversal.Fields("FhEntradaBodega"), "mm/dd/yyyy") & Chr(9) & strNumeroDoc & Chr(9) & strNumeroDoc & Chr(9) & rstUniversal.Fields("Cuenta") & Chr(9) & "Recaudo" & Chr(9) & "2" & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & "0" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "0"
          End If
          rstUniversal.MoveNext
        Loop
        CerrarRecorset rstUniversal
        

        
        'total c x p
        Dim strTipoCuenta As String
        strTipoCuenta = 2
        If Val(rstDespachosExp.Fields("VehiculoPropio")) = 0 Then
          douValor = rstDespachosExp.Fields("SaldoDesp")
          douValor = douValor - (rstDespachosExp.Fields("TRecaudo") + rstDespachosExp.Fields("TotalCE"))
        Else
          douValor = rstDespachosExp.Fields("TRecaudo") + rstDespachosExp.Fields("TotalCE")
          strTipoCuenta = 1
        End If

        If douValor < 0 Then
          douValor = douValor * -1
          strTipoCuenta = 1
        End If
        Print #1, "28150505" & Chr(9) & "00025" & Chr(9) & Format(rstDespachosExp.Fields("FhExpedicion"), "mm/dd/yyyy") & Chr(9) & strNumero & Chr(9) & strNumero & Chr(9) & strNit & Chr(9) & "Total c x p" & Chr(9) & strTipoCuenta & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & "0" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "0"
        rstDespachosExp.Close
        rstDespachosExp.Open "UPDATE despachos SET ExportadoContabilidad=1 WHERE OrdDespacho=" & LstDespachos.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
        LstDespachos.ListItems.Remove (II)
      Else
       II = II + 1
      End If
    Wend
    
    Close #1
  
  Exit Sub
'Error_Handler:
    'If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
  rstDespachosExp.CursorLocation = adUseClient
  VerDespachos
End Sub

Private Function Rellenar(Dato As String, Tama�o As Integer, Caracter As String, Posicion As Byte) As String
  FufuSt = ""
  If Len(Dato) < Tama�o Then
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

Private Sub Form_Unload(Cancel As Integer)
  Set rstDespachosExp.DataSource = Nothing
End Sub
