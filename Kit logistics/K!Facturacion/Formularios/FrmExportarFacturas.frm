VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmExportarFacturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar Facturas"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   14310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExportarSiigo 
      Caption         =   "Exportar SIIGO"
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "Consultar"
      Height          =   375
      Left            =   10800
      TabIndex        =   12
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Orden"
      Height          =   855
      Left            =   9000
      TabIndex        =   9
      Top             =   6240
      Width           =   1695
      Begin VB.OptionButton OptNumero 
         Caption         =   "Numero"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdExportarilimitadaContai 
      Caption         =   "Exportar Ilimitada Contai"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton CmdExportarTerceros 
      Caption         =   "Exportar terceros Altius"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton CmdExportarAltius 
      Caption         =   "Exportar Altius"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   6240
      Width           =   1815
   End
   Begin VB.TextBox TxtRuta 
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   6015
   End
   Begin VB.CommandButton CmdSeleccionarTodo 
      Caption         =   "Seleccionar todo"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   12480
      TabIndex        =   2
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton CmdExportar 
      Caption         =   "Exportar Ilimitada SCI"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   6240
      Width           =   2175
   End
   Begin MSComctlLib.ListView LstFacturas 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   13935
      _ExtentX        =   24580
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Factura"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tp"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tercero"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "VrFlete"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "VrManejo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "VrOtros"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label LblRuta 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   390
   End
End
Attribute VB_Name = "FrmExportarFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstFacturasExp As New ADODB.Recordset

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

Private Sub CmdConsultar_Click()
  VerFacturas
End Sub

Private Sub CmdExportarilimitadaContai_Click()
  Dim RutaSalida As String
  Dim Fila        As Long
  Dim Columna     As Long
  
On Error GoTo Error_Handler
    RutaSalida = TxtRuta.Text & "facexp" & Format(Date, "dd.mm.yy") & "." & Format(Time, "hh.mm.ss") & ".txt"
    Dim J As Integer
    Dim strCuenta As String
    Dim strDetalle As String
    Dim intTipo As Integer
    Dim douValor As Double
    Dim strNumero As String
    Dim intNroRegistros As Integer
    Fila = 0
    II = 1
    Open RutaSalida For Append As #1
    IniProg LstFacturas.ListItems.Count
    Print #1, "Cuenta  Comprobante Fecha(mm/dd/yyyy) Documento Documento Ref.  Nit Detalle Tipo  Valor Base  Centro de Costo Trans. Ext  Plazo"
    While II <= LstFacturas.ListItems.Count
      If LstFacturas.ListItems(II).Checked = True Then
        rstFacturasExp.Open "SELECT facturas_venta.*, terceros.RazonSocial " & _
                            "FROM facturas_venta " & _
                            "LEFT JOIN terceros ON facturas_venta.IdTercero = terceros.IdTercero " & _
                            "WHERE Exportada=0 AND Numero = " & LstFacturas.ListItems(II) & " AND TipoFactura = " & LstFacturas.ListItems(II).SubItems(1), CnnPrincipal, adOpenDynamic, adLockOptimistic
        'Corrientes o contados
        If Val(rstFacturasExp.Fields("TipoFactura")) = 1 Or Val(rstFacturasExp.Fields("TipoFactura")) = 2 Then
          intNroRegistros = 3
        ElseIf Val(rstFacturasExp.Fields("TipoFactura")) = 3 Then
          intNroRegistros = 4
        Else
          intNroRegistros = 0
        End If
        
        For J = 1 To intNroRegistros Step 1
          Fila = Fila + 1
          Select Case Val(rstFacturasExp.Fields("TipoFactura"))
            'Corriente
            Case 1
              Select Case J
                Case 1
                  strCuenta = "41450505"
                  intTipo = 2
                  strDetalle = "FLETES"
                  douValor = rstFacturasExp.Fields("VrFlete")
                Case 2
                  strCuenta = "41454005"
                  intTipo = 2
                  strDetalle = "Valor Seguro Docto"
                  douValor = rstFacturasExp.Fields("VrManejo")
                Case 3
                  strCuenta = "13050501"
                  intTipo = 1
                  strDetalle = "VLR TOTAL DOC"
                  douValor = rstFacturasExp.Fields("Total")
              End Select
              strNumero = Rellenar("B" & rstFacturasExp.Fields("Numero"), 9, "0", 1)
              
            'Contado
            Case 2
              Select Case J
                Case 1
                  strCuenta = "41450510"
                  intTipo = 2
                  strDetalle = "FLETES"
                  douValor = rstFacturasExp.Fields("VrFlete")
                Case 2
                  strCuenta = "41454005"
                  intTipo = 2
                  strDetalle = "Valor Seguro Docto"
                  douValor = rstFacturasExp.Fields("VrManejo")
                Case 3
                  strCuenta = "11050515"
                  intTipo = 1
                  strDetalle = "VLR TOTAL DOC"
                  douValor = rstFacturasExp.Fields("Total")
              End Select
              strNumero = Rellenar("A" & rstFacturasExp.Fields("Numero"), 9, "0", 1)
              
            'Destino
            Case 3
              Select Case J
                Case 1
                  strCuenta = "41450515"
                  intTipo = 2
                  strDetalle = "FLETES"
                  douValor = rstFacturasExp.Fields("VrFlete")
                Case 2
                  strCuenta = "41454005"
                  intTipo = 2
                  strDetalle = "Valor Seguro Docto"
                  douValor = rstFacturasExp.Fields("VrManejo")
                Case 3
                  strCuenta = "13050502"
                  intTipo = 1
                  strDetalle = "VLR TOTAL DOC"
                  douValor = rstFacturasExp.Fields("VrFlete")
                Case 4
                  strCuenta = "11050515"
                  intTipo = 1
                  strDetalle = "VLR TOTAL DOC"
                  douValor = rstFacturasExp.Fields("VrManejo")
              End Select
              strNumero = Rellenar("A" & rstFacturasExp.Fields("Numero"), 9, "0", 1)
          End Select
          
          Print #1, strCuenta & Chr(9) & "00003" & Chr(9) & Format(rstFacturasExp.Fields("Fecha"), "mm/dd/yyyy") & Chr(9) & strNumero & Chr(9) & strNumero & Chr(9) & rstFacturasExp.Fields("IdTercero") & Chr(9) & strDetalle & Chr(9) & intTipo & Chr(9) & Format(douValor, "##0.00;(##0.00)") & Chr(9) & "0" & Chr(9) & "404" & Chr(9) & "" & Chr(9) & "0"
          
        Next
          
        
        rstFacturasExp.Close
        rstFacturasExp.Open "UPDATE facturas_venta SET Exportada=1 where Numero=" & LstFacturas.ListItems(II) & " AND TipoFactura = " & LstFacturas.ListItems(II).SubItems(1), CnnPrincipal, adOpenDynamic, adLockOptimistic
        LstFacturas.ListItems.Remove (II)
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

Private Sub CmdExportar_Click()
  II = 1
  Open TxtRuta.Text & "facexp" & Format(Date, "dd.mm.yy") & "." & Format(Time, "hh.mm.ss") & ".txt" For Append As #1
  IniProg LstFacturas.ListItems.Count
  While II <= LstFacturas.ListItems.Count
    If LstFacturas.ListItems(II).Checked = True Then
      rstFacturasExp.Open "SELECT facturas_venta.*, terceros.RazonSocial " & _
                          "FROM facturas_venta " & _
                          "LEFT JOIN terceros ON facturas_venta.IdTercero = terceros.IdTercero " & _
                          "WHERE Exportada=0 AND Numero = " & LstFacturas.ListItems(II) & " AND TipoFactura = " & LstFacturas.ListItems(II).SubItems(1), CnnPrincipal, adOpenDynamic, adLockOptimistic
        Print #1, Format(rstFacturasExp.Fields("Fecha"), "yyyymmdd") & "30" & Rellenar(rstFacturasExp.Fields("Numero"), 9, "0", 1) & Rellenar(rstFacturasExp.Fields("IdTercero"), 11, " ", 1) & "99010001                                          41450505  001NNLMR" & Rellenar(rstFacturasExp.Fields("Numero"), 9, "0", 1) & Rellenar(rstFacturasExp.Fields("RazonSocial"), 50, " ", 2) & "0001                                    404              N      FLETES                        99"
        Print #1, "9999                 " & Rellenar(rstFacturasExp.Fields("IdTercero"), 11, " ", 1) & "00                            0             1.00" & Rellenar(Format(rstFacturasExp.Fields("VrFlete"), "##0.00;(##0.00)"), 17, " ", 1) & "  0.00  0.00             0.00             0.009999" & Rellenar(Format(rstFacturasExp.Fields("VrManejo"), "##0.00;(##0.00)"), 17, " ", 1) & "9999" & Rellenar(Format(rstFacturasExp.Fields("VrFlete"), "##0.00;(##0.00)"), 17, " ", 1) & "             0.00             0.00             0.00             1.00"
        Print #1, " N NNNN              0.0099             0.00                            99             0.00                0.00      000000305             0.00             0.00    99N99S  0             0.00             0.00             0.00" & Rellenar(Format(rstFacturasExp.Fields("VrFlete"), "##0.00;(##0.00)"), 17, " ", 1)
        Print #1, "             1.00                                                                    999999"
      rstFacturasExp.Close
      rstFacturasExp.Open "UPDATE facturas_venta SET Exportada=1 where Numero=" & LstFacturas.ListItems(II) & " AND TipoFactura = " & LstFacturas.ListItems(II).SubItems(1), CnnPrincipal, adOpenDynamic, adLockOptimistic
      LstFacturas.ListItems.Remove (II)
    Else
     II = II + 1
    End If
    Prog (II)
  Wend
  FinProg
  Close #1
  
  
End Sub

Private Sub CmdExportarAltius_Click()
  Dim rstFactura As New ADODB.Recordset
  rstFactura.CursorLocation = adUseClient
  Dim strSql As String
  II = 1
  Open TxtRuta.Text & "facexp" & Format(Date, "dd.mm.yy") & "." & Format(Time, "hh.mm.ss") & ".txt" For Append As #1
  IniProg LstFacturas.ListItems.Count
  While II <= LstFacturas.ListItems.Count
    If LstFacturas.ListItems(II).Checked = True Then
      strSql = "SELECT facturas_venta.*, terceros.RazonSocial, ciudades.CuentaFlete, ciudades.CuentaManejo, ciudades.CuentaCartera, Prefijo " & _
                          "FROM facturas_venta " & _
                          "LEFT JOIN terceros ON facturas_venta.IdTercero = terceros.IdTercero " & _
                          "LEFT JOIN ciudades ON terceros.IdCiudad = ciudades.IdCiudad " & _
                          "LEFT JOIN facturas_tipos ON facturas_venta.TipoFactura = facturas_tipos.IdTipoFactura " & _
                          "WHERE Exportada=0 AND Numero = " & LstFacturas.ListItems(II) & " AND TipoFactura = " & LstFacturas.ListItems(II).SubItems(1)
      rstFactura.Open strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
      'Cuenta Flete
      Print #1, Chr(34) & Format(Date, "dd") & Chr(34) & "," & Chr(34) & rstFactura!Prefijo & Chr(34) & "," & Chr(34) & rstFactura!Numero & Chr(34) & "," & Chr(34) & rstFactura!CuentaFlete & Chr(34) & "," & Replace(rstFactura!VrFlete, ",", ".") & "," & Chr(34) & "C" & Chr(34) & _
                  "," & Chr(34) & rstFactura!IDTercero & Chr(34) & "," & "0" & "," & "0" & "," & Chr(34) & Chr(34) & "," & Chr(34) & "Factura venta " & Chr(34) & _
                  "," & Chr(34) & " " & Chr(34) & "," & "0" & "," & "0" & "," & Chr(34) & rstFactura!Prefijo & "-" & rstFactura!Numero & Chr(34) & "," & Format(rstFactura!Fecha, "dd/mm/yyyy") & _
                  "," & Format(rstFactura!FhVence, "dd/mm/yyyy") & "," & Chr(34) & "C" & Chr(34)
      'Cuenta Manejo
      Print #1, Chr(34) & Format(Date, "dd") & Chr(34) & "," & Chr(34) & rstFactura!Prefijo & Chr(34) & "," & Chr(34) & rstFactura!Numero & Chr(34) & "," & Chr(34) & rstFactura!CuentaManejo & Chr(34) & ", " & Replace(rstFactura!VrManejo, ",", ".") & ", " & Chr(34) & "C" & Chr(34) & "" _
                  ; "," & Chr(34) & rstFactura!IDTercero & Chr(34) & "," & "0" & "," & "0" & "," & Chr(34) & Chr(34) & "," & Chr(34) & "Factura venta " & Chr(34) & _
                  "," & Chr(34) & " " & Chr(34) & "," & "0" & "," & "0" & "," & Chr(34) & rstFactura!Prefijo & "-" & rstFactura!Numero & Chr(34) & "," & Format(rstFactura!Fecha, "dd/mm/yyyy") & _
                  "," & Format(rstFactura!FhVence, "dd/mm/yyyy") & "," & Chr(34) & "C" & Chr(34)
      
      'Cuenta clientes nacionales (Cartera)
      Print #1, Chr(34) & Format(Date, "dd") & Chr(34) & "," & Chr(34) & rstFactura!Prefijo & Chr(34) & "," & Chr(34) & rstFactura!Numero & Chr(34) & "," & Chr(34) & rstFactura!CuentaCartera & Chr(34) & "," & Replace(rstFactura!VrFlete + rstFactura!VrManejo, ",", ".") & "," & Chr(34) & "D" & Chr(34) & _
                  "," & Chr(34) & rstFactura!IDTercero & Chr(34) & "," & "0" & "," & "0" & "," & Chr(34) & Chr(34) & "," & Chr(34) & "Factura venta " & Chr(34) & _
                  "," & Chr(34) & " " & Chr(34) & "," & "0" & "," & "0" & "," & Chr(34) & rstFactura!Prefijo & "-" & rstFactura!Numero & Chr(34) & "," & Format(rstFactura!Fecha, "dd/mm/yyyy") & _
                  "," & Format(rstFactura!FhVence, "dd/mm/yyyy") & "," & Chr(34) & "C" & Chr(34)
                  
        'Cree debito 13551525
        Print #1, Chr(34) & Format(Date, "dd") & Chr(34) & "," & Chr(34) & rstFactura!Prefijo & Chr(34) & "," & Chr(34) & rstFactura!Numero & Chr(34) & "," & Chr(34) & "13552001" & Chr(34) & "," & Replace((rstFactura!Total * 0.6 / 100), ",", ".") & "," & Chr(34) & "D" & Chr(34) & _
                    "," & Chr(34) & rstFactura!IDTercero & Chr(34) & "," & rstFactura!Total & "," & "0.6" & "," & Chr(34) & Chr(34) & "," & Chr(34) & "Factura venta " & Chr(34) & _
                    "," & Chr(34) & " " & Chr(34) & "," & "0" & "," & "0" & "," & Chr(34) & rstFactura!Prefijo & "-" & rstFactura!Numero & Chr(34) & "," & Format(rstFactura!Fecha, "dd/mm/yyyy") & _
                    "," & Format(rstFactura!FhVence, "dd/mm/yyyy") & "," & Chr(34) & "C" & Chr(34)
                    
        'Cree credito 23690510
        Print #1, Chr(34) & Format(Date, "dd") & Chr(34) & "," & Chr(34) & rstFactura!Prefijo & Chr(34) & "," & Chr(34) & rstFactura!Numero & Chr(34) & "," & Chr(34) & "23657501" & Chr(34) & "," & Replace((rstFactura!Total * 0.6 / 100), ",", ".") & "," & Chr(34) & "C" & Chr(34) & _
                    "," & Chr(34) & rstFactura!IDTercero & Chr(34) & "," & rstFactura!Total & "," & "0.6" & "," & Chr(34) & Chr(34) & "," & Chr(34) & "Factura venta " & Chr(34) & _
                    "," & Chr(34) & " " & Chr(34) & "," & "0" & "," & "0" & "," & Chr(34) & rstFactura!Prefijo & "-" & rstFactura!Numero & Chr(34) & "," & Format(rstFactura!Fecha, "dd/mm/yyyy") & _
                    "," & Format(rstFactura!FhVence, "dd/mm/yyyy") & "," & Chr(34) & "C" & Chr(34)

      rstFactura.Close
      rstFactura.Open "UPDATE facturas_venta SET Exportada=1 where Numero=" & LstFacturas.ListItems(II) & " AND TipoFactura = " & LstFacturas.ListItems(II).SubItems(1), CnnPrincipal, adOpenDynamic, adLockOptimistic
      LstFacturas.ListItems.Remove (II)
    Else
     II = II + 1
    End If
    Prog (II)
  Wend
  FinProg
  Close #1
End Sub




Private Sub CmdExportarSiigo_Click()
  Dim rstFactura As New ADODB.Recordset
  rstFactura.CursorLocation = adUseClient
  Dim strSql As String
  Dim intSecuencia As Integer
  II = 1
  intSecuencia = 1
  Open TxtRuta.Text & "facsiigoexp" & Format(Date, "dd.mm.yy") & "." & Format(Time, "hh.mm.ss") & ".txt" For Append As #1
  IniProg LstFacturas.ListItems.Count
  While II <= LstFacturas.ListItems.Count
    If LstFacturas.ListItems(II).Checked = True Then
      strSql = "SELECT facturas_venta.*, terceros.RazonSocial, ciudades.CuentaFlete, ciudades.CuentaManejo, ciudades.CuentaCartera, Prefijo " & _
                          "FROM facturas_venta " & _
                          "LEFT JOIN terceros ON facturas_venta.IdTercero = terceros.IdTercero " & _
                          "LEFT JOIN ciudades ON terceros.IdCiudad = ciudades.IdCiudad " & _
                          "LEFT JOIN facturas_tipos ON facturas_venta.TipoFactura = facturas_tipos.IdTipoFactura " & _
                          "WHERE Exportada=0 AND Numero = " & LstFacturas.ListItems(II) & " AND TipoFactura = " & LstFacturas.ListItems(II).SubItems(1)
      rstFactura.Open strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
      'Cuenta Flete
      Print #1, "F001" & Rellenar(rstFactura!Numero, 11, "0", 1) & Rellenar(intSecuencia, 5, "0", 1) & Rellenar(rstFactura!IDTercero, 13, "0", 1) & "000"
      intSecuencia = intSecuencia + 1

      rstFactura.Close
      'rstFactura.Open "UPDATE facturas_venta SET Exportada=1 where Numero=" & LstFacturas.ListItems(II) & " AND TipoFactura = " & LstFacturas.ListItems(II).SubItems(1), CnnPrincipal, adOpenDynamic, adLockOptimistic
      LstFacturas.ListItems.Remove (II)
    Else
     II = II + 1
    End If
    Prog (II)
  Wend
  FinProg
  Close #1
End Sub

Private Sub CmdExportarTerceros_Click()
  Dim rstTerceros As New ADODB.Recordset
  rstTerceros.CursorLocation = adUseClient
  Dim strSql As String
  strSql = "SELECT terceros.*, ciudades.IdDepartamento, ciudades.NmCiudad, ciudades.CodigoDivision, departamentos.NmDepartamento " & _
           "FROM terceros " & _
           "LEFT JOIN ciudades ON terceros.IdCiudad = ciudades.IdCiudad " & _
           "LEFT JOIN departamentos ON ciudades.IdDepartamento = departamentos.IdDepartamento"
           
  AbrirRecorset rstTerceros, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
  Open TxtRuta.Text & "terexp" & Format(Date, "dd.mm.yy") & "." & Format(Time, "hh.mm.ss") & ".txt" For Append As #1
  While rstTerceros.EOF = False
      Print #1, Chr(34) & Chr(34) & "," & Chr(34) & rstTerceros!TpDoc & Chr(34) & "," & Chr(34) & rstTerceros!IDTercero & Chr(34) & "," & Chr(34) & DigitoVerificacion(rstTerceros!IDTercero) & Chr(34) & ","; Chr(34) & rstTerceros!Apellido1 & Chr(34) & "," & Chr(34) & rstTerceros!Apellido2 & Chr(34) & "," & Chr(34) & rstTerceros!Nombre & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & rstTerceros!RazonSocial & Chr(34) & "," & Chr(34) & rstTerceros!RazonSocial & Chr(34) & "," & Chr(34) & rstTerceros!Direccion & Chr(34) & "," & Chr(34) & rstTerceros!Telefono & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & rstTerceros!Celular & Chr(34) & "," & Chr(34) & rstTerceros!IdDepartamento & Chr(34) & "," & Chr(34) & rstTerceros!NmDepartamento & Chr(34) & "," & Chr(34) & rstTerceros!CodigoDivision & Chr(34) & "," & Chr(34) & _
      rstTerceros!NmCiudad & Chr(34) & "," & Chr(34) & "0" & Chr(34) & "," & Chr(34) & "01" & Chr(34) & "," & Chr(34) & "0" & Chr(34) & "," & Chr(34) & "0" & Chr(34) & "," & Chr(34) & "0" & Chr(34) & _
      "," & Chr(34) & "V00001" & Chr(34) & "," & Chr(34) & "N" & Chr(34) & "," & Chr(34) & "0" & Chr(34) & "," & Chr(34) & "" & Chr(34) & "," & Chr(34) & "" & Chr(34) & "," & Chr(34) & "" & Chr(34) & "," & Chr(34) & "" & Chr(34) & "," & Chr(34) & "00" & Chr(34) & "," & Chr(34) & rstTerceros!Email & Chr(34) & "," & Chr(34) & "" & Chr(34) & "," & Chr(34) & "" & Chr(34) & "," & Chr(34) & "" & Chr(34) & "," & Chr(34) & rstTerceros!CodigoDivision & Chr(34) & "," & Chr(34) & "" & Chr(34) & "," & Chr(34) & "" & Chr(34) & "," & Chr(34) & "" & Chr(34) & "," & Chr(34) & "" & Chr(34) & "," & Chr(34) & "" & Chr(34) & "," & Chr(34) & "" & Chr(34) & "," & Chr(34) & "" & Chr(34)
    rstTerceros.MoveNext
  Wend
  Close #1
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub
Private Sub VerFacturas()
  LstFacturas.ListItems.Clear
  rstFacturasExp.Open "SELECT facturas_venta.*, terceros.RazonSocial, NmTipoFactura " & _
                          "FROM facturas_venta " & _
                          "LEFT JOIN terceros ON facturas_venta.IdTercero = terceros.IdTercero " & _
                          "LEFT JOIN facturas_tipos ON facturas_venta.TipoFactura = facturas_tipos.IdTipoFactura " & _
                          "WHERE Exportada=0 " & _
                          DevOrden, CnnPrincipal, adOpenDynamic, adLockOptimistic
  IniProg rstFacturasExp.RecordCount
  If rstFacturasExp.RecordCount > 0 Then
    Do While rstFacturasExp.EOF = False
      Prog (rstFacturasExp.AbsolutePosition)
      Set Item = LstFacturas.ListItems.Add(, , rstFacturasExp!Numero)
      Item.SubItems(1) = rstFacturasExp!TipoFactura
      Item.SubItems(2) = rstFacturasExp!NmTipoFactura
      Item.SubItems(3) = Format(rstFacturasExp!Fecha, "dd/mm/yy")
      Item.SubItems(4) = rstFacturasExp!RazonSocial & ""
      Item.SubItems(5) = Format(rstFacturasExp!VrFlete, "0;(0)")
      Item.SubItems(6) = Format(rstFacturasExp!VrManejo, "0;(0)")
      Item.SubItems(7) = Format(rstFacturasExp!VrOtros, "0;(0)")
      Item.SubItems(8) = Format(rstFacturasExp!Total, "0;(0)")
      rstFacturasExp.MoveNext
    Loop
  End If
  FinProg
  rstFacturasExp.Close
End Sub

Private Function DevOrden() As String
  If OptTipo.Value = True Then
    DevOrden = "ORDER BY TipoFactura, Numero"
  Else
    DevOrden = "ORDER BY Numero, TipoFactura"
  End If
End Function


Private Sub CmdSeleccionarTodo_Click()
  II = 1
  For II = 1 To LstFacturas.ListItems.Count
    LstFacturas.ListItems(II).Checked = True
  Next
End Sub


Private Sub Form_Load()
  rstFacturasExp.CursorLocation = adUseClient
  TxtRuta.Text = GetSetting("Kit Logistics", "Facturacion", "RutaExportarArchivoFacturas")
  VerFacturas
End Sub
