VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCuentasCobrar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas cobrar"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdGenerarReciboPorGuia 
      Caption         =   "Seleccionar recibo de caja por guia"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   7560
      Width           =   3495
   End
   Begin VB.CommandButton CmdExportarExcel 
      Caption         =   "Exportar excel"
      Height          =   375
      Left            =   9720
      TabIndex        =   3
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CommandButton CmdGenerarReciboCaja 
      Caption         =   "Generar recibo de caja seleccionados"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7560
      Width           =   3495
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   11880
      TabIndex        =   0
      Top             =   7080
      Width           =   1815
   End
   Begin MSComctlLib.ListView LstCuentasCobrar 
      Height          =   6855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   12091
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipo"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Numero"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fecha"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Cliente"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "D.V"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Saldo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Centro Operaciones"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Soporte"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComCtl2.DTPicker DPFechaPago 
      Height          =   300
      Left            =   1080
      TabIndex        =   5
      Top             =   7080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      Format          =   49872897
      CurrentDate     =   38971
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Fecha pago:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   7080
      Width           =   900
   End
End
Attribute VB_Name = "FrmCuentasCobrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCuentasCobrar As New ADODB.Recordset

Private Sub CmdExportarExcel_Click()
  AbrirRecorset rstUniversal, "SELECT sql_ic_cartera_edades.* from sql_ic_cartera_edades where TipoFactura = 2 OR TipoFactura = 3 Order by TipoFactura, FechaDoc, RazonSocial", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.State = adStateOpen Then
    If rstUniversal.EOF = False Then
      ExportarExcel rstUniversal
    End If
  End If
  CerrarRecorset rstUniversal
End Sub

Private Sub CmdGenerarReciboCaja_Click()
  Dim rstCuenta As New ADODB.Recordset
  rstCuenta.CursorLocation = adUseClient
  Dim rstRecibo As New ADODB.Recordset
  rstRecibo.CursorLocation = adUseClient
  II = 1
    While II <= LstCuentasCobrar.ListItems.Count
      If LstCuentasCobrar.ListItems(II).Checked = True Then
        FufuSt = "Select cuentas_cobrar.* from cuentas_cobrar where IdCxC=" & LstCuentasCobrar.ListItems.Item(II)
        AbrirRecorset rstCuenta, FufuSt, CnnPrincipal, adOpenDynamic, adLockReadOnly
        FufuLo = SacarConsecutivo("RecibosCaja", CnnPrincipal)
        AbrirRecorset rstUniversal, "INSERT INTO recibos_caja (numero, Fecha, IdTercero, Total, Comentarios, codigo_banco_fk, Impreso, IdReciboTipo, FechaPago) " & _
                                        "VALUES (" & FufuLo & ",'" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', '" & rstCuenta.Fields("IdTercero") & "', " & rstCuenta.Fields("Saldo") & ", '', 6, 1, 2, '" & Format(DPFechaPago.value, "yyyy-mm-dd") & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
        FufuSt = "Select recibos_caja.* from recibos_caja where numero=" & FufuLo
        AbrirRecorset rstRecibo, FufuSt, CnnPrincipal, adOpenDynamic, adLockReadOnly
        
        AbrirRecorset rstUniversal, "INSERT INTO recibos_caja_det (IdRecibo, codigo_cuenta_cobrar_fk, valor) VALUES (" & rstRecibo.Fields("IdRecibo") & ", " & rstCuenta.Fields("IdCxC") & ", " & rstCuenta.Fields("Saldo") & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
        AbrirRecorset rstUniversal, "UPDATE cuentas_cobrar SET Saldo = Saldo - " & rstCuenta.Fields("Saldo") & " WHERE IdCxC = " & rstCuenta.Fields("IdCxC"), CnnPrincipal, adOpenDynamic, adLockOptimistic
        LstCuentasCobrar.ListItems.Remove (II)
      Else
       II = II + 1
      End If
    Wend
End Sub

Private Sub CmdGenerarReciboPorGuia_Click()
Dim rstCuentaCobrarTemp As New ADODB.Recordset
Dim rstActualizar As New ADODB.Recordset
Dim rstRecibo As New ADODB.Recordset
Dim numeroFactura As Long
Dim tipoFactura As Integer
rstCuentaCobrarTemp.CursorLocation = adUseClient
rstActualizar.CursorLocation = adUseClient
rstRecibo.CursorLocation = adUseClient
  If Principal.ToolConsultas1.AbrirDevDatos("Digite el numero de la guia", "Digite el numero de la guia que desea generar recibo", 3, 0) = True Then
    AbrirRecorset rstUniversal, "Select Guia, GuiaTipo from guias where Guia=" & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstUniversal.RecordCount > 0 Then
      numeroFactura = rstUniversal.Fields("Guia")
      tipoFactura = rstUniversal.Fields("GuiaTipo")
      If tipoFactura = 2 Or tipoFactura = 3 Then
        AbrirRecorset rstCuentaCobrarTemp, "Select IdCxC, Saldo, IdTercero from cuentas_cobrar where TipoFactura=" & tipoFactura & " AND NroDocumento = " & numeroFactura, CnnPrincipal, adOpenDynamic, adLockOptimistic
        If rstCuentaCobrarTemp.RecordCount > 0 Then
          If Val(rstCuentaCobrarTemp.Fields("Saldo")) > 0 Then
            FufuLo = SacarConsecutivo("RecibosCaja", CnnPrincipal)
            AbrirRecorset rstActualizar, "INSERT INTO recibos_caja (numero, Fecha, IdTercero, Total, Comentarios, codigo_banco_fk, Impreso, IdReciboTipo, FechaPago) " & _
                                            "VALUES (" & FufuLo & ",'" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', '" & rstCuentaCobrarTemp.Fields("IdTercero") & "', " & rstCuentaCobrarTemp.Fields("Saldo") & ", '', 6, 1, 2, '" & Format(DPFechaPago.value, "yyyy-mm-dd") & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
            FufuSt = "Select recibos_caja.* from recibos_caja where numero=" & FufuLo
            AbrirRecorset rstRecibo, FufuSt, CnnPrincipal, adOpenDynamic, adLockReadOnly
            AbrirRecorset rstActualizar, "INSERT INTO recibos_caja_det (IdRecibo, codigo_cuenta_cobrar_fk, valor) VALUES (" & rstRecibo.Fields("IdRecibo") & ", " & rstCuentaCobrarTemp.Fields("IdCxC") & ", " & rstCuentaCobrarTemp.Fields("Saldo") & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
            AbrirRecorset rstActualizar, "UPDATE cuentas_cobrar SET Saldo = Saldo - " & rstCuentaCobrarTemp.Fields("Saldo") & " WHERE IdCxC = " & rstCuentaCobrarTemp.Fields("IdCxC"), CnnPrincipal, adOpenDynamic, adLockOptimistic
            CerrarRecorset rstRecibo
            Ver
          Else
            MsgBox "la cuenta no tiene saldo"
          End If
        Else
          MsgBox "No existe la cuenta por cobrar, verifique si la guia ya fue exportada a contabilidad"
        End If
        CerrarRecorset rstCuentaCobrarTemp
      Else
        MsgBox "Solo se puede realizar recibo automatico a guias contado-destino"
      End If
      CmdGenerarReciboPorGuia_Click
    Else
      MsgBox "No se encontro la guia", vbCritical
      CmdGenerarReciboPorGuia_Click
    End If
    CerrarRecorset rstUniversal
  End If
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  rstCuentasCobrar.CursorLocation = adUseClient
  DPFechaPago.value = Date
  Ver
End Sub

Private Sub Ver()
  LstCuentasCobrar.ListItems.Clear
  AbrirRecorset rstUniversal, "SELECT sql_ic_cartera_edades.* from sql_ic_cartera_edades where TipoFactura = 2 OR TipoFactura = 3 Order by TipoFactura, FechaDoc, RazonSocial", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Do While rstUniversal.EOF = False
      Set Item = LstCuentasCobrar.ListItems.Add(, , rstUniversal.Fields("IdCxC"))
      Item.SubItems(1) = rstUniversal.Fields("NmTipoFactura")
      Item.SubItems(2) = rstUniversal.Fields("NroDocumento")
      Item.SubItems(3) = rstUniversal.Fields("FechaDoc")
      Item.SubItems(4) = rstUniversal.Fields("RazonSocial")
      Item.SubItems(5) = rstUniversal.Fields("DiasVencida")
      Item.SubItems(6) = rstUniversal.Fields("Total")
      Item.SubItems(7) = rstUniversal.Fields("Saldo")
      Item.SubItems(8) = rstUniversal.Fields("NmPuntoOperaciones") & ""
      Item.SubItems(9) = rstUniversal.Fields("Soporte") & ""
      rstUniversal.MoveNext
    Loop
  CerrarRecorset rstUniversal
End Sub
