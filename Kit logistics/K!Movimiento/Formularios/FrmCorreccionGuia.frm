VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCorreccionGuia 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Correccion de guias..."
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtRemitente 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   6615
   End
   Begin VB.TextBox TxtKilosRealesAnt 
      Height          =   285
      Left            =   2520
      TabIndex        =   47
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox TxtKilosVolumenAnt 
      Height          =   285
      Left            =   2520
      TabIndex        =   46
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox TxtKilosFacturadosAnt 
      Height          =   285
      Left            =   2520
      TabIndex        =   45
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox TxtUnidadesAnt 
      Height          =   285
      Left            =   2520
      TabIndex        =   44
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox TxtFleteAnt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6360
      TabIndex        =   43
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox TxtManejoAnt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6360
      TabIndex        =   42
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox TxtDeclaradoAnt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6360
      TabIndex        =   41
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox TxtRecaudoAnt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6360
      TabIndex        =   40
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox TxtIdNegociacion 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox TxtNmNegociacion 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   38
      Top             =   840
      Width           =   5295
   End
   Begin VB.TextBox TxtIdCiudadOrigen 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox TxtNmOrigen 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   36
      Top             =   1560
      Width           =   5295
   End
   Begin VB.ComboBox CboTpServicio 
      Height          =   315
      ItemData        =   "FrmCorreccionGuia.frx":0000
      Left            =   1080
      List            =   "FrmCorreccionGuia.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox TxtRecaudo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4920
      TabIndex        =   14
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox TxtNmDestino 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   31
      Top             =   1920
      Width           =   5295
   End
   Begin VB.TextBox TxtDocCliente 
      Height          =   285
      Left            =   4920
      TabIndex        =   5
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox TxtIdCiudadDestino 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox TxtNmTercero 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   29
      Top             =   480
      Width           =   5295
   End
   Begin VB.TextBox TxtCuenta 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox TxtComentarios 
      Height          =   765
      Left            =   1080
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   5640
      Width           =   6615
   End
   Begin VB.TextBox TxtUnidades 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox TxtKilosFacturados 
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox TxtKilosVolumen 
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox TxtKilosReales 
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1080
      TabIndex        =   23
      Top             =   6480
      Width           =   1455
   End
   Begin VB.TextBox TxtDeclarado 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4920
      TabIndex        =   11
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox TxtManejo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4920
      TabIndex        =   13
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox TxtFlete 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4920
      TabIndex        =   12
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox TxtGuia 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   18
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   6240
      TabIndex        =   16
      Top             =   6480
      Width           =   1455
   End
   Begin MSComctlLib.ListView LstGuiasTipos 
      Height          =   1455
      Left            =   1080
      TabIndex        =   35
      Top             =   4080
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Remitente:"
      Height          =   195
      Left            =   120
      TabIndex        =   48
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Negociacion:"
      Height          =   195
      Left            =   120
      TabIndex        =   39
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Origen:"
      Height          =   195
      Left            =   480
      TabIndex        =   37
      Top             =   1560
      Width           =   510
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Recaudo:"
      Height          =   195
      Left            =   4080
      TabIndex        =   34
      Top             =   3720
      Width           =   705
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Documento:"
      Height          =   195
      Left            =   3960
      TabIndex        =   33
      Top             =   2280
      Width           =   870
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Destino:"
      Height          =   195
      Left            =   360
      TabIndex        =   32
      Top             =   1920
      Width           =   585
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Tercero:"
      Height          =   195
      Left            =   360
      TabIndex        =   30
      Top             =   480
      Width           =   600
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Comentarios:"
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   5640
      Width           =   915
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Unidades:"
      Height          =   195
      Left            =   240
      TabIndex        =   27
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "K Fac:"
      Height          =   195
      Left            =   600
      TabIndex        =   26
      Top             =   3720
      Width           =   465
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "K Vol:"
      Height          =   195
      Left            =   540
      TabIndex        =   25
      Top             =   3360
      Width           =   420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "K Reales:"
      Height          =   195
      Left            =   240
      TabIndex        =   24
      Top             =   3000
      Width           =   690
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Manejo:"
      Height          =   195
      Left            =   4320
      TabIndex        =   22
      Top             =   3360
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Flete:"
      Height          =   195
      Left            =   4440
      TabIndex        =   21
      Top             =   3000
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Declarado:"
      Height          =   195
      Left            =   4080
      TabIndex        =   20
      Top             =   2640
      Width           =   780
   End
   Begin VB.Label LblTitulos 
      AutoSize        =   -1  'True
      Caption         =   "Tp Servicio:"
      Height          =   195
      Index           =   9
      Left            =   105
      TabIndex        =   19
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Guia:"
      Height          =   195
      Left            =   585
      TabIndex        =   17
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "FrmCorreccionGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstGuia As New ADODB.Recordset
Dim rstCorreccion As New ADODB.Recordset



Dim liquidado As Boolean
Dim PermiteRecaudo As Boolean
Dim NegociacionInactiva As Boolean
Dim ListaPreciosVencida As Boolean
Dim rstGuias As New ADODB.Recordset
Dim strSqlGuias As String
Dim douPorcentajeManejo As Double
Dim douMinimoManejoUnidad As Double
Dim douMinimoManejoDespacho As Double
Dim intIdListaPrecios As Integer
Dim douVrKilo As Double
Dim douDctoKilo As Double
Dim douKilosMinimos As Double
Dim boolNoAplicarDctoReexpediciones As Integer
Dim boolRedondearFlete As Integer
Private Sub CmdAceptar_Click()
Dim intTipoCobro As Integer
Dim intGuiaFactura As Integer
  AbrirRecorset rstUniversal, "SELECT guias_tipos.* FROM guias_tipos WHERE IdGuiaTipo = " & LstGuiasTipos.SelectedItem, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.RecordCount > 0 Then
    intTipoCobro = rstUniversal!TipoCobro
    intGuiaFactura = rstUniversal!GuiaFactura
  End If
  CerrarRecorset rstUniversal
  
  AbrirRecorset rstCorreccion, "Update Guias set GuiFac = " & intGuiaFactura & ", VrDeclarado=" & Val(TxtDeclarado) & ", VrFlete=" & Val(TxtFlete.Text) & ", VrManejo=" & Val(TxtManejo.Text) & ", TpServicio=" & CboTpServicio.ListIndex & ", KilosReales = " & Val(TxtKilosReales.Text) & ", KilosVolumen = " & Val(TxtKilosVolumen.Text) & ", KilosFacturados = " & Val(TxtKilosFacturados.Text) & ", Unidades = " & Val(TxtUnidades.Text) & ", GuiaTipo=" & LstGuiasTipos.SelectedItem & ", TipoCobro = " & intTipoCobro & ", Cliente='" & TxtNmTercero.Text & "', Cuenta='" & TxtCuenta.Text & "', IdCiuDestino=" & Val(TxtIdCiudadDestino.Text) & ", IdCiuOrigen=" & Val(TxtIdCiudadOrigen.Text) & ", DocCliente='" & TxtDocCliente.Text & "', Remitente='" & TxtRemitente.Text & "', Recaudo=" & Val(TxtRecaudo.Text) & "  where Guia=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
  InsertarLog 8, Val(TxtGuia.Text)
  CerrarRecorset rstCorreccion
  AbrirRecorset rstCorreccion, "Insert into Correcciones (GuiaCorregida, FechaCorreccion, CuentaC, IdUsuarioCorreccion, IdTpServicio, VrDeclaradoC, VrFleteC, VrManejoC, GuiaFacC, KilosRealesC, KilosVolumenC, KilosFacturadosC, UnidadesC, GuiaTipoC, Comentarios, IdCiuDestinoC, DocClienteC, RecaudoC)" & _
  " values (" & FufuLo & ",'" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "','" & TxtCuenta.Text & "'," & CodUsuarioActivo & ", " & Val(rstGuia!TpServicio) & ", " & rstGuia.Fields("VrDeclarado") & ", " & rstGuia.Fields("VrFlete") & ", " & rstGuia.Fields("VrManejo") & "," & rstGuia!GuiFac & "," & rstGuia!kilosReales & "," & rstGuia!KilosVolumen & "," & rstGuia!KilosFacturados & "," & rstGuia!unidades & ", " & rstGuia!GuiaTipo & ", '" & TxtComentarios.Text & "'," & Val(TxtIdCiudadDestino.Text) & ", '" & TxtDocCliente.Text & "', " & Val(TxtRecaudo.Text) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
  MsgBox "La guia a sido corregida satisfactoriamente", vbInformation
  Unload Me
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub
Private Sub Form_Load()
  rstGuia.CursorLocation = adUseClient
  rstCorreccion.CursorLocation = adUseClient
  
  rstGuia.Open "Select guias.* from guias where guia=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstGuia.RecordCount > 0 Then
    TxtGuia.Text = FufuLo
    TxtDeclaradoAnt.Text = rstGuia.Fields("VrDeclarado")
    TxtFleteAnt.Text = rstGuia.Fields("VrFlete")
    TxtManejoAnt.Text = rstGuia.Fields("VrManejo")
    TxtUnidadesAnt.Text = rstGuia.Fields("Unidades")
    TxtKilosRealesAnt.Text = rstGuia.Fields("KilosReales")
    TxtKilosVolumenAnt.Text = rstGuia.Fields("KilosVolumen")
    TxtKilosFacturadosAnt.Text = rstGuia.Fields("KilosFacturados")
    CboTpServicio.ListIndex = Val(rstGuia!TpServicio)
    TxtCuenta.Text = rstGuia.Fields("Cuenta")
    TxtNmTercero.Text = rstGuia.Fields("Cliente")
    TxtIdCiudadDestino.Text = rstGuia.Fields("IdCiuDestino")
    If Val(TxtIdCiudadDestino.Text) <> 0 Then TxtNmDestino.Text = DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Val(TxtIdCiudadDestino.Text), "NmCiudad", CnnPrincipal)
    
    TxtIdCiudadOrigen.Text = rstGuia.Fields("IdCiuOrigen")
    If Val(TxtIdCiudadOrigen.Text) <> 0 Then TxtNmOrigen.Text = DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Val(TxtIdCiudadOrigen.Text), "NmCiudad", CnnPrincipal)
    
    TxtDocCliente.Text = rstGuia.Fields("DocCliente") & ""
    TxtRemitente.Text = rstGuia.Fields("Remitente") & ""
    TxtRecaudoAnt.Text = rstGuia.Fields("Recaudo")
    Dim rstGuiasTipos As New ADODB.Recordset
    rstGuiasTipos.CursorLocation = adUseClient
    AbrirRecorset rstGuiasTipos, "SELECT guias_tipos.* from guias_tipos WHERE 1", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      Do While rstGuiasTipos.EOF = False
        Set Item = LstGuiasTipos.ListItems.Add(, , rstGuiasTipos!IdGuiaTipo)
        Item.SubItems(1) = rstGuiasTipos!NmGuiaTipo
        rstGuiasTipos.MoveNext
      Loop
    CerrarRecorset rstGuiasTipos
    Set Item = LstGuiasTipos.FindItem(rstGuia!GuiaTipo)
    If Item Is Nothing Then
      MsgBox "No se encontro el tipo de guia"
    Else
      LstGuiasTipos.ListItems(Item.Index).Selected = True
    End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstGuia = Nothing
  Set rstCorreccion = Nothing
End Sub

Private Sub TxtComentarios_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys vbTab
  End If
End Sub

Private Sub TxtCuenta_GotFocus()
  EnfocarT TxtCuenta
End Sub

Private Sub TxtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsulta 7, CnnPrincipal
    TxtCuenta.Text = Principal.ToolConsultas1.DatSt
  End If
End Sub

Private Sub TxtCuenta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtCuenta_Validate(Cancel As Boolean)
  AbrirRecorset rstUniversal, "Select IdTercero, RazonSocial, IdCliente from Terceros where IdTercero='" & TxtCuenta.Text & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = False Then
    TxtNmTercero.Text = rstUniversal.Fields("RazonSocial") & ""
  Else
    TxtNmTercero.Text = ""
  End If
  CerrarRecorset rstUniversal
  If TxtCuenta.Text <> "" Then
    AbrirRecorset rstUniversal, "Select IdTercero, RazonSocial, IdCliente from Terceros where IdTercero='" & TxtCuenta & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        TxtIdNegociacion.Text = rstUniversal.Fields("IdCliente") & ""
        CerrarRecorset rstUniversal
        CargarNegociacion
      Else
          TxtCuenta.Text = ""
      End If
    CerrarRecorset rstUniversal
  End If
  
End Sub

Private Sub TxtDeclarado_GotFocus()
EnfocarT TxtDeclarado
End Sub

Private Sub TxtDeclarado_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtDeclarado_Validate(Cancel As Boolean)
  If Val(TxtManejo.Text) = 0 And Val(TxtDeclarado.Text) > 0 Then
    TxtManejo.Text = Val(TxtDeclarado.Text) * douPorcentajeManejo / 100
    If douMinimoManejoDespacho > Val(TxtManejo.Text) Then
      TxtManejo.Text = douMinimoManejoDespacho
    End If
    If (douMinimoManejoUnidad * Val(TxtUnidades.Text)) > Val(TxtManejo.Text) Then
      TxtManejo.Text = douMinimoManejoUnidad * Val(TxtUnidades.Text)
    End If
  End If
End Sub

Private Sub TxtDocCliente_GotFocus()
EnfocarT TxtDocCliente
End Sub

Private Sub TxtDocCliente_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtFlete_GotFocus()
EnfocarT TxtFlete
End Sub

Private Sub TxtFlete_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdCiudadDestino_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsulta 1, CnnPrincipal
    TxtIdCiudadDestino.Text = Principal.ToolConsultas1.DatLo
  End If
End Sub

Private Sub TxtIdCiudadDestino_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdCiudadDestino_Validate(Cancel As Boolean)
  AbrirRecorset rstUniversal, "SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Val(TxtIdCiudadDestino.Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = False Then
    TxtNmDestino.Text = rstUniversal.Fields("NmCiudad")
  Else
    TxtNmDestino.Text = ""
  End If
  CerrarRecorset rstUniversal
End Sub

Private Sub TxtIdCiudadOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsulta 1, CnnPrincipal
    TxtIdCiudadOrigen.Text = Principal.ToolConsultas1.DatLo
  End If
End Sub

Private Sub TxtIdCiudadOrigen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdCiudadOrigen_Validate(Cancel As Boolean)
  AbrirRecorset rstUniversal, "SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Val(TxtIdCiudadOrigen.Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = False Then
    TxtNmOrigen.Text = rstUniversal.Fields("NmCiudad")
  Else
    TxtNmOrigen.Text = ""
  End If
  CerrarRecorset rstUniversal
End Sub

Private Sub TxtIdNegociacion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtKilosFacturados_GotFocus()
EnfocarT TxtKilosFacturados
End Sub

Private Sub TxtKilosFacturados_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtKilosFacturados_Validate(Cancel As Boolean)
  If Val(TxtFlete.Text) = 0 And Val(TxtIdCiudadDestino.Text) <> 0 Then
    Dim rstListaPreciosDetalle As New ADODB.Recordset
    rstListaPreciosDetalle.CursorLocation = adUseClient
    AbrirRecorset rstListaPreciosDetalle, "SELECT VrKilo, Minimos FROM listaspreciosciudades WHERE IdListaPrecios = " & intIdListaPrecios & " AND IdCiudadOrigen = " & Val(TxtIdCiudadOrigen.Text) & " AND IdCiudad = " & Val(TxtIdCiudadDestino.Text) & " AND IdProducto = 1 AND VrKilo > 0", CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstListaPreciosDetalle.RecordCount > 0 Then
      douVrKilo = rstListaPreciosDetalle!VrKilo
      AbrirRecorset rstUniversal, "SELECT Reexpedicion FROM ciudades WHERE Reexpedicion = 1 AND IdCiudad = " & Val(TxtIdCiudadDestino.Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.RecordCount > 0 Then
        If boolNoAplicarDctoReexpediciones = 1 Then
          douDctoKilo = 0
        End If
      End If
      Dim douFlete As Double
      douFlete = Val(TxtKilosFacturados.Text) * (douVrKilo - (douVrKilo * douDctoKilo / 100))
      If boolRedondearFlete = 1 Then
        douFlete = Round(douFlete * 0.01) * 100
      Else
        douFlete = Round(douFlete)
      End If
      TxtFlete.Text = douFlete
    End If
  End If
End Sub

Private Sub TxtKilosReales_GotFocus()
EnfocarT TxtKilosReales
End Sub

Private Sub TxtKilosReales_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtKilosReales_Validate(Cancel As Boolean)
  LiquidarKilosFacturar
End Sub

Private Sub TxtKilosVolumen_GotFocus()
EnfocarT TxtKilosVolumen
End Sub

Private Sub TxtKilosVolumen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtKilosVolumen_Validate(Cancel As Boolean)
  LiquidarKilosFacturar
End Sub

Private Sub TxtManejo_GotFocus()
EnfocarT TxtManejo
End Sub

Private Sub TxtManejo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtRecaudo_GotFocus()
  EnfocarT TxtRecaudo
End Sub

Private Sub TxtUnidades_GotFocus()
EnfocarT TxtUnidades
End Sub

Private Sub TxtUnidades_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtUnidades_Validate(Cancel As Boolean)
      AbrirRecorset rstUniversal, "SELECT Minimos FROM listaspreciosciudades WHERE IdListaPrecios = " & intIdListaPrecios & " AND IdCiudadOrigen = " & Val(TxtIdCiudadOrigen.Text) & " AND IdCiudad = " & Val(TxtIdCiudadDestino.Text) & " AND IdProducto = 1 AND VrKilo > 0", CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.RecordCount > 0 Then
        If Val(TxtKilosFacturados.Text) <= 0 Then
          If Val(rstUniversal!Minimos) > 0 Then
            TxtKilosFacturados.Text = Val(rstUniversal!Minimos) * Val(TxtUnidades.Text)
          End If
        End If
      End If
      If douKilosMinimos > 0 Then
        TxtKilosFacturados.Text = douKilosMinimos * Val(TxtUnidades.Text)
        If Val(TxtKilosReales.Text) <= 0 Then
          TxtKilosReales.Text = douKilosMinimos * Val(TxtUnidades.Text)
        End If
      End If
      
      CerrarRecorset rstUniversal
End Sub


Private Sub CargarNegociacion()
  Dim rstListaPrecios As New ADODB.Recordset
  rstListaPrecios.CursorLocation = adUseClient
  
  AbrirRecorset rstUniversal, "SELECT negociaciones.* FROM negociaciones WHERE Id=" & Val(TxtIdNegociacion), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = False Then
      douPorcentajeManejo = rstUniversal!PorManejo
      douMinimoManejoUnidad = rstUniversal!MinUniManejo
      douMinimoManejoDespacho = rstUniversal!MinDesManejo
      douKilosMinimos = rstUniversal!Minimos
      douDctoKilo = rstUniversal!DctoK
      boolNoAplicarDctoReexpediciones = DevCheck(rstUniversal!NoAplicarDctoReexpediciones)
      PermiteRecaudo = rstUniversal!PermiteRecaudo
      boolRedondearFlete = rstUniversal!RedondearFlete
      NegociacionInactiva = rstUniversal!Inactivo
      intIdListaPrecios = rstUniversal!ListaPrecios
      TxtNmNegociacion.Text = rstUniversal!NmNegociacion
  Else
    TxtNmNegociacion.Text = "": TxtIdNegociacion.Text = "0"
  End If
  CerrarRecorset rstUniversal
End Sub


Private Sub LiquidarKilosFacturar()
  If Val(TxtKilosReales.Text) > Val(TxtKilosFacturados.Text) Then
    TxtKilosFacturados.Text = Val(TxtKilosReales.Text)
  End If
  If Val(TxtKilosVolumen.Text) > Val(TxtKilosFacturados.Text) Then
    TxtKilosFacturados.Text = Val(TxtKilosVolumen.Text)
  End If
End Sub
