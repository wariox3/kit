VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCorreccionGuia 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Correccion de guias..."
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtIdCiudadOrigen 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox TxtNmOrigen 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   34
      Top             =   840
      Width           =   3735
   End
   Begin VB.ComboBox CboTpServicio 
      Height          =   315
      ItemData        =   "FrmCorreccionGuia.frx":0000
      Left            =   1080
      List            =   "FrmCorreccionGuia.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox TxtRecaudo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3360
      TabIndex        =   31
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox TxtNmDestino 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   28
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox TxtDocCliente 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox TxtIdCiudadDestino 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox TxtNmTercero 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   26
      Top             =   480
      Width           =   3735
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
      TabIndex        =   12
      Top             =   5280
      Width           =   5055
   End
   Begin VB.TextBox TxtUnidades 
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox TxtKilosFacturados 
      Height          =   285
      Left            =   3360
      TabIndex        =   11
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox TxtKilosVolumen 
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox TxtKilosReales 
      Height          =   285
      Left            =   3360
      TabIndex        =   9
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1080
      TabIndex        =   20
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox TxtDeclarado 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox TxtManejo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox TxtFlete 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   2640
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
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4680
      TabIndex        =   13
      Top             =   6120
      Width           =   1455
   End
   Begin MSComctlLib.ListView LstGuiasTipos 
      Height          =   1455
      Left            =   1080
      TabIndex        =   33
      Top             =   3720
      Width           =   5055
      _ExtentX        =   8916
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
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Origen:"
      Height          =   195
      Left            =   480
      TabIndex        =   35
      Top             =   840
      Width           =   510
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Recaudo:"
      Height          =   195
      Left            =   2520
      TabIndex        =   32
      Top             =   2280
      Width           =   705
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Documento:"
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   1560
      Width           =   870
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Destino:"
      Height          =   195
      Left            =   360
      TabIndex        =   29
      Top             =   1200
      Width           =   585
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Tercero:"
      Height          =   195
      Left            =   360
      TabIndex        =   27
      Top             =   480
      Width           =   600
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Comentarios:"
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   5280
      Width           =   915
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Unidades:"
      Height          =   195
      Left            =   240
      TabIndex        =   24
      Top             =   3000
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "K Fac:"
      Height          =   195
      Left            =   2880
      TabIndex        =   23
      Top             =   3360
      Width           =   465
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "K Vol:"
      Height          =   195
      Left            =   540
      TabIndex        =   22
      Top             =   3360
      Width           =   420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "K Reales:"
      Height          =   195
      Left            =   2640
      TabIndex        =   21
      Top             =   3000
      Width           =   690
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Manejo:"
      Height          =   195
      Left            =   2640
      TabIndex        =   19
      Top             =   2640
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Flete:"
      Height          =   195
      Left            =   570
      TabIndex        =   18
      Top             =   2640
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Declarado:"
      Height          =   195
      Left            =   180
      TabIndex        =   17
      Top             =   2280
      Width           =   780
   End
   Begin VB.Label LblTitulos 
      AutoSize        =   -1  'True
      Caption         =   "Tp Servicio:"
      Height          =   195
      Index           =   9
      Left            =   105
      TabIndex        =   16
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Guia:"
      Height          =   195
      Left            =   585
      TabIndex        =   14
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

Private Sub CmdAceptar_Click()
Dim intTipoCobro As Integer
Dim intGuiaFactura As Integer
  AbrirRecorset rstUniversal, "SELECT guias_tipos.* FROM guias_tipos WHERE IdGuiaTipo = " & LstGuiasTipos.SelectedItem, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.RecordCount > 0 Then
    intTipoCobro = rstUniversal!TipoCobro
    intGuiaFactura = rstUniversal!GuiaFactura
  End If
  CerrarRecorset rstUniversal
  
  AbrirRecorset rstCorreccion, "Update Guias set GuiFac = " & intGuiaFactura & ", VrDeclarado=" & Val(TxtDeclarado) & ", VrFlete=" & Val(TxtFlete.Text) & ", VrManejo=" & Val(TxtManejo.Text) & ", TpServicio=" & CboTpServicio.ListIndex & ", KilosReales = " & Val(TxtKilosReales.Text) & ", KilosVolumen = " & Val(TxtKilosVolumen.Text) & ", KilosFacturados = " & Val(TxtKilosFacturados.Text) & ", Unidades = " & Val(TxtUnidades.Text) & ", GuiaTipo=" & LstGuiasTipos.SelectedItem & ", TipoCobro = " & intTipoCobro & ", Cliente='" & TxtNmTercero.Text & "', Cuenta='" & TxtCuenta.Text & "', IdCiuDestino=" & Val(TxtIdCiudadDestino.Text) & ", IdCiuOrigen=" & Val(TxtIdCiudadOrigen.Text) & ", DocCliente='" & TxtDocCliente.Text & "', Recaudo=" & Val(TxtRecaudo.Text) & "  where Guia=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
  InsertarLog 8, Val(TxtGuia.Text)
  CerrarRecorset rstCorreccion
  AbrirRecorset rstCorreccion, "Insert into Correcciones (GuiaCorregida, FechaCorreccion, CuentaC, IdUsuarioCorreccion, IdTpServicio, VrDeclaradoC, VrFleteC, VrManejoC, GuiaFacC, KilosRealesC, KilosVolumenC, KilosFacturadosC, UnidadesC, GuiaTipoC, Comentarios, IdCiuDestinoC, DocClienteC, RecaudoC)" & _
  " values (" & FufuLo & ",'" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "','" & TxtCuenta.Text & "'," & CodUsuarioActivo & ", " & Val(rstGuia!TpServicio) & ", " & rstGuia.Fields("VrDeclarado") & ", " & rstGuia.Fields("VrFlete") & ", " & rstGuia.Fields("VrManejo") & "," & rstGuia!GuiFac & "," & rstGuia!KilosReales & "," & rstGuia!KilosVolumen & "," & rstGuia!KilosFacturados & "," & rstGuia!Unidades & ", " & rstGuia!GuiaTipo & ", '" & TxtComentarios.Text & "'," & Val(TxtIdCiudadDestino.Text) & ", '" & TxtDocCliente.Text & "', " & Val(TxtRecaudo.Text) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
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
    TxtDeclarado.Text = rstGuia.Fields("VrDeclarado")
    TxtFlete.Text = rstGuia.Fields("VrFlete")
    TxtManejo.Text = rstGuia.Fields("VrManejo")
    TxtUnidades.Text = rstGuia.Fields("Unidades")
    TxtKilosReales.Text = rstGuia.Fields("KilosReales")
    TxtKilosVolumen.Text = rstGuia.Fields("KilosVolumen")
    TxtKilosFacturados.Text = rstGuia.Fields("KilosFacturados")
    CboTpServicio.ListIndex = Val(rstGuia!TpServicio)
    TxtCuenta.Text = rstGuia.Fields("Cuenta")
    TxtNmTercero.Text = rstGuia.Fields("Cliente")
    TxtIdCiudadDestino.Text = rstGuia.Fields("IdCiuDestino")
    If Val(TxtIdCiudadDestino.Text) <> 0 Then TxtNmDestino.Text = DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Val(TxtIdCiudadDestino.Text), "NmCiudad", CnnPrincipal)
    
    TxtIdCiudadOrigen.Text = rstGuia.Fields("IdCiuOrigen")
    If Val(TxtIdCiudadOrigen.Text) <> 0 Then TxtNmOrigen.Text = DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Val(TxtIdCiudadOrigen.Text), "NmCiudad", CnnPrincipal)
    
    TxtDocCliente.Text = rstGuia.Fields("DocCliente") & ""
    TxtRecaudo.Text = rstGuia.Fields("Recaudo")
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
End Sub

Private Sub TxtDeclarado_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
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

Private Sub TxtKilosFacturados_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtKilosReales_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtKilosVolumen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtManejo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtUnidades_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
