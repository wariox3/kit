VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmImportarBufalo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar guias bufalo"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSeleccionarTodo 
      Caption         =   "Seleccionar todo"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   9120
      TabIndex        =   5
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton CmdCrearGuias 
      Caption         =   "Crear guias"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton CmdCargar 
      Caption         =   "Cargar"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox TxtDespacho 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.ListView LstGuias 
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9551
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Guia"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Errores"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Despacho:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "FrmImportarBufalo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CnnBufalo As ADODB.Connection
Public rstDespacho As ADODB.Recordset
Public rstGuias As ADODB.Recordset
Private Sub CmdCargar_Click()
AbrirRecorset rstDespacho, "SELECT codigo_despacho_pk FROM tte_despacho where codigo_despacho_pk = " & Val(TxtDespacho.Text), CnnBufalo, adOpenDynamic, adLockOptimistic
If rstDespacho.RecordCount > 0 Then
  AbrirRecorset rstGuias, "SELECT codigo_guia_pk, consecutivo FROM tte_guia where codigo_despacho_proveedor_fk =" & Val(TxtDespacho.Text), CnnBufalo, adOpenDynamic, adLockOptimistic
  If rstGuias.RecordCount > 0 Then
    Do While rstGuias.EOF = False
      Set Item = LstGuias.ListItems.Add(, , rstGuias.Fields("codigo_guia_pk"))
      Item.SubItems(1) = rstGuias.Fields("consecutivo")
      rstGuias.MoveNext
    Loop
  End If
  CmdCargar.Enabled = False
Else
  MsgBox "El despacho no existe"
End If
End Sub

Private Sub CmdCrearGuias_Click()
  Dim rstGuia As New ADODB.Recordset
  rstGuia.CursorLocation = adUseClient
  
  Dim strSql As String
  Dim Guia As Double
  Dim flete As Double
  Dim declarado As Double
  Dim manejo As Double
  Dim unidades As Double
  Dim kilosFacturar As Integer
  Dim kilosReales As Integer
  Dim kilosVolumen As Integer
  Dim observaciones As String
  Dim Documento As String
  Dim orden As Integer
  II = 1
  While II <= LstGuias.ListItems.Count
    If LstGuias.ListItems(II).Checked = True Then
      Guia = LstGuias.ListItems(II).SubItems(1)
      strSql = "select Guia from guias where Guia = " & Guia
      If ExRecorset(strSql) = False Then
        strSql = "select codigo_guia_pk, tte_guia.codigo_ciudad_origen_fk, tte_guia.codigo_ciudad_destino_fk, tte_guia.devolver_documento, tte_guia.destinatario, tte_guia.direccion, tte_guia.telefono, tte_guia.flete, tte_guia.manejo, tte_guia.declarado, tte_guia.peso, tte_guia.peso_volumen, tte_guia.peso_facturar, tte_guia.cantidad, tte_guia.observacion, tte_guia.documento, tte_empresa.nit, tte_empresa.cuenta_kit, tte_empresa.nombre as empresa_nombre FROM tte_guia " & _
        "LEFT JOIN tte_empresa ON tte_guia.codigo_empresa_fk = tte_empresa.codigo_empresa_pk " & _
        "LEFT JOIN tte_ciudad as tte_ciudad_destino ON tte_guia.codigo_ciudad_destino_fk = tte_ciudad_destino.codigo_ciudad_pk " & _
        "LEFT JOIN tte_ciudad as tte_ciudad_origen ON tte_guia.codigo_ciudad_origen_fk = tte_ciudad_origen.codigo_ciudad_pk " & _
        "WHERE codigo_guia_pk = " & LstGuias.ListItems(II)
        AbrirRecorset rstGuia, strSql, CnnBufalo, adOpenDynamic, adLockOptimistic
        If rstGuia.RecordCount > 0 Then
            AbrirRecorset rstUniversal, "Select* from Rutas_Ciudades where IdCiudad=" & rstGuia.Fields("codigo_ciudad_destino_fk"), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
            If rstUniversal.EOF = False Then
              orden = rstUniversal.Fields("Orden")
            Else
              orden = 0
            End If
            flete = rstGuia.Fields("flete")
            manejo = rstGuia.Fields("manejo")
            declarado = rstGuia.Fields("declarado")
            kilosFacturar = rstGuia.Fields("peso_facturar")
            kilosReales = rstGuia.Fields("peso")
            kilosVolumen = rstGuia.Fields("peso_volumen")
            unidades = rstGuia.Fields("cantidad")
            Documento = rstGuia.Fields("documento") & ""
            observaciones = rstGuia.Fields("observacion") & ""
            strSql = "INSERT INTO Guias " & _
            "(Guia, CR, Remitente, IdCliente, DocCliente, NmDestinatario, DirDestinatario, TelDestinatario, IdCiuDestino, IdRuta, " & _
            "FhEntradaBodega, VrDeclarado, VrFlete, VrManejo, Unidades, KilosReales, KilosFacturados, KilosVolumen, " & _
            "Estado, IdFactura, IdDespacho, Observaciones, COIng, Cuenta, Cliente, Recaudo, orden, EmpaqueRef, RelCliente, IdCiuOrigen, TpServicio, CPorte, Entregada, Descargada, Despachada, Anulada, GuiFac, Facturada, IdUsuario, IdEmpresa, GuiaTipo, TipoCobro) " & _
            "VALUES(" & Guia & "," & Coperaciones & ",'" & rstGuia.Fields("empresa_nombre") & "', " & rstGuia.Fields("cuenta_kit") & ",'" & Documento & "','" & rstGuia.Fields("destinatario") & "','" & rstGuia.Fields("direccion") & "','" & Mid(rstGuia.Fields("telefono"), 1, 11) & "', " & rstGuia.Fields("codigo_ciudad_destino_fk") & ", 1, " & _
            "'" & Format(Date, "yyyy-mm-dd") & " " & Format(Time, "HH:mm:ss") & "', " & declarado & ", " & flete & ", " & manejo & ", " & unidades & ", " & kilosReales & ", " & kilosFacturar & ", " & kilosReales & ", " & _
            "'D', 0, null, '" & observaciones & "', " & Coperaciones & ", '" & rstGuia.Fields("nit") & "', '" & rstGuia.Fields("empresa_nombre") & "', 0, " & orden & ", '',''," & rstGuia.Fields("codigo_ciudad_origen_fk") & ", 0, " & rstGuia.Fields("devolver_documento") & ", 0, 0, 0, 0, 0, 0, " & CodUsuarioActivo & ", 1, 1, 3)"
            
            AbrirRecorset rstUniversal, strSql, CnnPrincipal, adOpenDynamic, adLockBatchOptimistic
            LstGuias.ListItems.Remove (II)
        Else
          II = II + 1
        End If
      Else
        LstGuias.ListItems(II).SubItems(2) = "La guia ya existe"
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

Private Sub CmdSeleccionarTodo_Click()
  II = 1
  For II = 1 To LstGuias.ListItems.Count
    LstGuias.ListItems(II).Checked = True
  Next
End Sub

Private Sub Form_Load()
  Set CnnBufalo = New ADODB.Connection
  Set rstDespacho = New ADODB.Recordset
  Set rstGuias = New ADODB.Recordset
  CnnBufalo.CursorLocation = adUseClient
  rstDespacho.CursorLocation = adUseClient
  rstGuias.CursorLocation = adUseClient
  CnnBufalo.Open "DRIVER=" & GetSetting("Kit Logistics", "Configuracion", "CnnDriver") & "; SERVER=" & GetSetting("Kit Logistics", "Configuracion", "CnnServidorBufalo") & "; PORT=5800; DATABASE=bdbufalo; PWD=tMq32.*++; UID=soporte;OPTION=3"
  
End Sub
