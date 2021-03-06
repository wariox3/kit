VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmImportarGuias2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar guias"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   12420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtFecha 
      Height          =   285
      Left            =   5760
      TabIndex        =   18
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton CmdSeleccionarTodo 
      Caption         =   "Seleccionar todo"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox TxtManejoMinimoUnidad 
      Height          =   285
      Left            =   7560
      TabIndex        =   14
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox TxtMininoPorDespacho 
      Height          =   285
      Left            =   5760
      TabIndex        =   12
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox TxtPorManejo 
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Text            =   "0.5"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox TxtDcto 
      Height          =   285
      Left            =   3480
      TabIndex        =   8
      Text            =   "35"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox TxtMinimos 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Text            =   "30"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox TxtIdTercero 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Text            =   "860001965"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton CmdCrearGuias 
      Caption         =   "Crear guias"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton CmdCargar 
      Caption         =   "Cargar"
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10440
      TabIndex        =   0
      Top             =   7080
      Width           =   1935
   End
   Begin MSComctlLib.ListView LstGuias 
      Height          =   6015
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   10610
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
      NumItems        =   18
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Guia"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CodCliente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Destinatario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Direccion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Telefono"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "CodCiudad"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Ciudad"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Departamento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Origen"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Und"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Peso"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Declarado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Observacion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Fact"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Flete"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Manejo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "K.Fact"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   195
      Left            =   5160
      TabIndex        =   17
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Min unidad:"
      Height          =   195
      Left            =   6600
      TabIndex        =   15
      Top             =   480
      Width           =   825
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Min despacho:"
      Height          =   195
      Left            =   4560
      TabIndex        =   13
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label Label5 
      Caption         =   "% Mjo:"
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Dcto:"
      Height          =   195
      Left            =   3000
      TabIndex        =   9
      Top             =   120
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Minimos:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   525
   End
End
Attribute VB_Name = "FrmImportarGuias2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdCrearGuias_Click()
  Dim rstCiudad As New ADODB.Recordset
  rstCiudad.CursorLocation = adUseClient
  Dim strSql As String
  Dim Guia As Double
  Dim flete As Double
  Dim declarado As Double
  Dim manejo As Double
  Dim unidades As Double
  Dim kilosFacturar As Integer
  Dim kilosReales As Integer
  Dim observaciones As String
  Dim DocumentoCliente As String
  
  Dim orden As Integer
  II = 1
  While II <= LstGuias.ListItems.Count
    If LstGuias.ListItems(II).Checked = True Then
      'MsgBox LstGuias.ListItems(II).SubItems(1)
      Guia = Val(LstGuias.ListItems(II).SubItems(1))
      strSql = "select Guia from guias where Guia = " & Guia
      If ExRecorset(strSql) = False Then
        strSql = "select IdCiudad  from ciudades where IdCiudad = " & LstGuias.ListItems(II).SubItems(6)
        AbrirRecorset rstCiudad, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
        If rstCiudad.RecordCount > 0 Then
          AbrirRecorset rstUniversal, "Select* from Rutas_Ciudades where IdCiudad=" & rstCiudad.Fields("IdCiudad"), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
          If rstUniversal.EOF = False Then
            orden = rstUniversal.Fields("Orden")
          Else
            orden = 0
          End If
          CerrarRecorset rstUniversal
          flete = LstGuias.ListItems(II).SubItems(15)
          manejo = LstGuias.ListItems(II).SubItems(16)
          declarado = Val(LstGuias.ListItems(II).SubItems(12))
          kilosFacturar = Val(LstGuias.ListItems(II).SubItems(17))
          kilosReales = Val(LstGuias.ListItems(II).SubItems(11))
          kilosReales = Round(kilosReales)
          unidades = LstGuias.ListItems(II).SubItems(10)
          observaciones = Mid(LstGuias.ListItems(II).SubItems(13), 1, 200)
          DocumentoCliente = LstGuias.ListItems(II).SubItems(2)
          
          strSql = "INSERT INTO Guias " & _
          "(Guia, CR, Remitente, IdCliente, DocCliente, NmDestinatario, DirDestinatario, TelDestinatario, IdCiuDestino, IdRuta, " & _
          "FhEntradaBodega, VrDeclarado, VrFlete, VrManejo, Unidades, KilosReales, KilosFacturados, KilosVolumen, " & _
          "Estado, IdFactura, IdDespacho, Observaciones, COIng, Cuenta, Cliente, Recaudo, orden, EmpaqueRef, RelCliente, IdCiuOrigen, TpServicio, CPorte, Entregada, Descargada, Despachada, Anulada, GuiFac, Facturada, IdUsuario, IdEmpresa, GuiaTipo, TipoCobro) " & _
          "VALUES(" & Guia & ",6,'TEXTILES LAFAYETTE S.A.S', '434','" & DocumentoCliente & "','" & LstGuias.ListItems(II).SubItems(3) & "','" & LstGuias.ListItems(II).SubItems(4) & "','" & LstGuias.ListItems(II).SubItems(5) & "', " & Val(LstGuias.ListItems(II).SubItems(6)) & ", 1, " & _
          "'" & TxtFecha.Text & "', " & declarado & ", " & flete & ", " & manejo & ", " & unidades & ", " & kilosReales & ", " & kilosFacturar & ", " & kilosReales & ", " & _
          "'I', 0, null, '" & observaciones & "', 6, '" & TxtIdTercero.Text & "', 'TEXTILES LAFAYETTE S.A.S', 0, " & orden & ", '','',20516, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 3)"
          AbrirRecorset rstUniversal, strSql, CnnPrincipal, adOpenDynamic, adLockBatchOptimistic
          AbrirRecorset rstUniversal, "DELETE FROM guias_importar WHERE Id = " & LstGuias.ListItems(II), CnnPrincipal, adOpenDynamic, adLockBatchOptimistic
        End If
        CerrarRecorset rstCiudad
        LstGuias.ListItems.Remove (II)
      Else
        II = II + 1
      End If
    Else
     II = II + 1
    End If
  Wend
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Function DevFlete(ciudadDestino As String, unidades As Integer, peso As Integer, boolRedondearFlete As Integer) As Double
  Dim rstListaPreciosDetalle As New ADODB.Recordset
  rstListaPreciosDetalle.CursorLocation = adUseClient
  Dim douVrKilo As Double
  
  AbrirRecorset rstListaPreciosDetalle, "SELECT VrKilo, Minimos FROM listaspreciosciudades WHERE IdListaPrecios = 1 AND IdCiudadOrigen = 20516 AND IdCiudad = " & Val(ciudadDestino) & " AND IdProducto = 1 AND VrKilo > 0", CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstListaPreciosDetalle.RecordCount > 0 Then
     douVrKilo = rstListaPreciosDetalle!VrKilo
     Dim douFlete As Double
     If unidades * Val(TxtMinimos) > peso Then
      peso = unidades * Val(TxtMinimos)
     End If
     douFlete = peso * douVrKilo
     If Val(TxtDcto.Text) > 0 Then
      douFlete = douFlete - (douFlete * Val(TxtDcto.Text) / 100)
     End If
     If boolRedondearFlete = 1 Then
       douFlete = Round(douFlete * 0.01) * 100
     Else
       douFlete = Round(douFlete)
     End If
     DevFlete = douFlete
   End If
End Function

Private Function DevManejo(declarado As Double, unidades As Integer) As Double
  Dim douPorcentajeManejo As Double
  Dim douMinimoManejoUnidad As Double
  Dim douMinimoManejoDespacho As Double
  Dim manejo As Double
  
  douPorcentajeManejo = Val(TxtPorManejo.Text)
  douMinimoManejoUnidad = Val(TxtManejoMinimoUnidad.Text)
  douMinimoManejoDespacho = Val(TxtMininoPorDespacho)
  manejo = declarado * douPorcentajeManejo / 100
  If douMinimoManejoDespacho > manejo Then
    manejo = douMinimoManejoDespacho
  End If
  If douMinimoManejoUnidad * unidades > manejo Then
    manejo = douMinimoManejoUnidad * unidades
  End If
  DevManejo = manejo
End Function

Private Function DevKilosFacturar(unidades As Integer, peso As Integer) As Integer
     If unidades * Val(TxtMinimos) > peso Then
      peso = unidades * Val(TxtMinimos)
     End If
     DevKilosFacturar = peso
End Function

Private Sub CmdSeleccionarTodo_Click()
  II = 1
  For II = 1 To LstGuias.ListItems.Count
    LstGuias.ListItems(II).Checked = True
  Next
End Sub

Private Sub Form_Load()
  TxtFecha.Text = Format(Date, "yyyy/mm/dd") & " " & Format(Time, "HH:mm:ss")
  'AbrirRecorset rstUniversal, "SELECT configuracion.consecutivo_importar_guia FROM configuracion", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  '  TxtConsecutivo.Text = rstUniversal.Fields("consecutivo_importar_guia")
  'CerrarRecorset rstUniversal
  AbrirRecorset rstUniversal, "SELECT MinDesManejo FROM negociaciones WHERE Id= 434", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    TxtMininoPorDespacho.Text = rstUniversal.Fields("MinDesManejo")
  CerrarRecorset rstUniversal
  
  cargar
End Sub

Sub cargar()
  Dim flete As Double
  Dim manejo As Double
  Dim kilosFacturar As Integer
    
  AbrirRecorset rstUniversal, "SELECT guias_importar.* FROM guias_importar", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
      flete = DevFlete(rstUniversal!codigoCiudad, rstUniversal!unidades, rstUniversal!peso, 0)
      manejo = DevManejo(rstUniversal!declarado, rstUniversal!unidades)
      kilosFacturar = DevKilosFacturar(rstUniversal!unidades, rstUniversal!peso)
      Set Item = LstGuias.ListItems.Add(, , rstUniversal!Id)
        Item.SubItems(1) = rstUniversal!Guia & ""
        Item.SubItems(2) = rstUniversal!DocumentoCliente & ""
        'Destinatario
        Item.SubItems(3) = rstUniversal!NmDestinatario & ""
        'Direccion
        Item.SubItems(4) = rstUniversal!DirDestinatario & ""
        Item.SubItems(5) = rstUniversal!TelDestinatario & ""
        Item.SubItems(6) = rstUniversal!codigoCiudad
        Item.SubItems(7) = ""
        Item.SubItems(8) = ""
        Item.SubItems(9) = ""
        'Unidades
        Item.SubItems(10) = rstUniversal!unidades
        Item.SubItems(11) = rstUniversal!peso
        Item.SubItems(12) = rstUniversal!declarado
        Item.SubItems(13) = ""
        Item.SubItems(14) = ""
        'Flete
        Item.SubItems(15) = flete
        Item.SubItems(16) = manejo
        Item.SubItems(17) = kilosFacturar
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
End Sub
