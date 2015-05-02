VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmPendientesPorFacturar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pendientes por facturar por cuenta..."
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12465
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   12465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Ordenado"
      Height          =   855
      Left            =   3960
      TabIndex        =   14
      Top             =   5760
      Width           =   1695
      Begin VB.OptionButton OptFecha 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton OptGuia 
         Caption         =   "Guia"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdGuiaFactura 
      Caption         =   "Buscar guias factura"
      Height          =   255
      Left            =   9840
      TabIndex        =   13
      Top             =   6240
      Width           =   2535
   End
   Begin VB.CheckBox ChkConNegociacion 
      Caption         =   "Con negociacion"
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton CmdBuscarDocumento 
      Caption         =   "Buscar documento"
      Height          =   255
      Left            =   6240
      TabIndex        =   10
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton CmdBuscarGuia 
      Caption         =   "Buscar guia"
      Height          =   255
      Left            =   8040
      TabIndex        =   9
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton CmdCambiarNegociacion 
      Caption         =   "Cambiar"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox TxtIdNegociacion 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton CmdFiltrar 
      Caption         =   "&Filtrar"
      Height          =   255
      Left            =   9840
      TabIndex        =   5
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   255
      Left            =   11280
      TabIndex        =   4
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox TxtId 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox TxtNombre 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   5400
      Width           =   7095
   End
   Begin MSDataGridLib.DataGrid GrillaGuias 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   9128
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
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "Guia"
         Caption         =   "Guia"
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
         DataField       =   "FhEntradaBodega"
         Caption         =   "Fecha"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "DocCliente"
         Caption         =   "Documento"
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
         DataField       =   "NmCiudad"
         Caption         =   "Destino"
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
         DataField       =   "NmDestinatario"
         Caption         =   "Destinatario"
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
         DataField       =   "VrDeclarado"
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
      BeginProperty Column06 
         DataField       =   "VrManejo"
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
      BeginProperty Column07 
         DataField       =   "VrFlete"
         Caption         =   "Flete"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Unidades"
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
      BeginProperty Column09 
         DataField       =   "KilosReales"
         Caption         =   "KR"
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
      BeginProperty Column10 
         DataField       =   "KilosFacturados"
         Caption         =   "KF"
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
      BeginProperty Column11 
         DataField       =   "KilosVolumen"
         Caption         =   "KV"
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
      BeginProperty Column12 
         DataField       =   "GuiFac"
         Caption         =   "GF"
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
      BeginProperty Column13 
         DataField       =   "Remitente"
         Caption         =   "Remitente"
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
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   2475.213
         EndProperty
      EndProperty
   End
   Begin VB.Label LblNroRegistros 
      Caption         =   "O registros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   9840
      TabIndex        =   12
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Negociacion:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   945
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   3
      Top             =   5400
      Width           =   210
   End
End
Attribute VB_Name = "FrmPendientesPorFacturar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstTemGuias As New ADODB.Recordset

Private Sub CmdBuscarDocumento_Click()
  If Principal.ToolConsultas1.AbrirDevDatos("Numero de documento", "Digite el numero del documento", 2, 0) = True Then
    If rstTemGuias.State = adStateOpen Then Set rstTemGuias = Nothing
    AbrirRecorset rstTemGuias, "Select*from guias where Facturada=0 and Anulada=0 and GuiFac=0 and DocCliente='" & Val(Principal.ToolConsultas1.DatSt) & "' and (IdTpCtaFlete=3 or IdTpCtaManejo=3)", CnnPrincipal, adOpenDynamic, adLockOptimistic
    Set GrillaGuias.DataSource = rstTemGuias
    LblNroRegistros.Caption = rstTemGuias.RecordCount & " Registros encontrados"
  End If
End Sub

Private Sub CmdBuscarGuia_Click()
  If Principal.ToolConsultas1.AbrirDevDatos("Numero de Guia", "Digite el numero de la guia que desea buscar", 3, 0) = True Then
    If rstTemGuias.State = adStateOpen Then Set rstTemGuias = Nothing
    AbrirRecorset rstTemGuias, "Select*from guias where Facturada=0 and Anulada=0 and GuiFac=0 and Guia=" & Val(Principal.ToolConsultas1.DatLo) & " and (IdTpCtaFlete=3 or IdTpCtaManejo=3)", CnnPrincipal, adOpenDynamic, adLockOptimistic
    Set GrillaGuias.DataSource = rstTemGuias
    LblNroRegistros.Caption = rstTemGuias.RecordCount & " Registros encontrados"
  End If
End Sub

Private Sub CmdCambiarNegociacion_Click()
  If TxtId.Text <> "" Then
    FufuSt = TxtId
    FrmBuscarNegociaciones.Show 1
    If FufuLo <> 0 Then
      TxtIdNegociacion.Text = FufuLo
    End If
  End If
End Sub

Private Sub CmdFiltrar_Click()
  If ChkConNegociacion.Value = 1 Then
    If TxtId.Text <> "" And Val(TxtIdNegociacion) <> 0 Then
      AbrirRecorset rstTemGuias, "Select guias.*, Ciudades.NmCiudad from guias left join Ciudades on guias.idciudestino=ciudades.idciudad where Facturada=0 and Anulada=0 and GuiFac=0 and Cuenta='" & TxtId.Text & "' and IdCliente=" & Val(TxtIdNegociacion.Text) & " and TipoCobro = 3 " & OrdenamientoGrilla, CnnPrincipal, adOpenDynamic, adLockOptimistic
      Set GrillaGuias.DataSource = rstTemGuias
      LblNroRegistros.Caption = rstTemGuias.RecordCount & " Registros encontrados"
    Else
      MsgBox "Debe especificar una negociacion y una cuenta", vbCritical
    End If
  Else
    If TxtId.Text <> "" Then
      AbrirRecorset rstTemGuias, "Select guias.*, Ciudades.NmCiudad from guias left join Ciudades on guias.idciudestino=ciudades.idciudad where Facturada=0 and Anulada=0 and GuiFac=0 and Cuenta='" & TxtId.Text & "' AND TipoCobro = 3 " & OrdenamientoGrilla, CnnPrincipal, adOpenDynamic, adLockOptimistic
      
      Set GrillaGuias.DataSource = rstTemGuias
      LblNroRegistros.Caption = rstTemGuias.RecordCount & " Registros encontrados"
    Else
      MsgBox "Debe especificar una cuenta", vbCritical
    End If
  End If
End Sub

Private Sub CmdGuiaFactura_Click()
  AbrirRecorset rstTemGuias, "Select guias.*, Ciudades.NmCiudad from guias, Ciudades where (guias.idciudestino=ciudades.idciudad) and Facturada=0 and Anulada=0 and GuiFac=1 and Cuenta='" & TxtId.Text & "' and (IdTpCtaFlete=3 or IdTpCtaManejo=3) order by FhEntradaBodega", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Set GrillaGuias.DataSource = rstTemGuias
  LblNroRegistros.Caption = rstTemGuias.RecordCount & " Registros encontrados"
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub


Private Sub Form_Load()
  rstTemGuias.CursorLocation = adUseClient
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstTemGuias = Nothing
End Sub

Private Sub GrillaGuias_BeforeUpdate(Cancel As Integer)
  AbrirRecorset rstUniversal, "Insert into Correcciones (GuiaCorregida, FechaCorreccion, CuentaC, IdUsuarioCorreccion, IdTpServicio, VrDeclaradoC, VrFleteC, VrManejoC, GuiaFacC, KilosRealesC, KilosVolumenC, KilosFacturadosC, UnidadesC, IdTpCtaFleteC, IdTpCtaManejoC, Comentarios) values (" & rstTemGuias.Fields("Guia").OriginalValue & ",'" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "','" & rstTemGuias.Fields("Cuenta").OriginalValue & "'," & CodUsuarioActivo & ", " & Val(rstTemGuias.Fields("TpServicio").OriginalValue) & ", " & rstTemGuias.Fields("VrDeclarado").OriginalValue & ", " & rstTemGuias.Fields("VrFlete").OriginalValue & ", " & rstTemGuias.Fields("VrManejo").OriginalValue & "," & rstTemGuias.Fields("GuiFac").OriginalValue & "," & rstTemGuias.Fields("KilosReales").OriginalValue & "," & rstTemGuias.Fields("KilosVolumen").OriginalValue & _
  "," & rstTemGuias.Fields("KilosFacturados").OriginalValue & "," & rstTemGuias.Fields("Unidades").OriginalValue & ", " & rstTemGuias.Fields("IdTpCtaFlete").OriginalValue & _
  ", " & rstTemGuias.Fields("IdTpCtaManejo").OriginalValue & ", 'Modificado desde facturacion')", CnnPrincipal, adOpenDynamic, adLockOptimistic
End Sub
Private Sub TxtId_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsulta 7, CnnPrincipal
    TxtId.Text = Principal.ToolConsultas1.DatSt
  End If
End Sub
Private Sub TxtId_Validate(Cancel As Boolean)
  If TxtId.Text <> "" Then
    AbrirRecorset rstUniversal, "Select IdTercero, RazonSocial, IdCliente from Terceros where IdTercero='" & TxtId.Text & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      TxtNombre = rstUniversal!RazonSocial & ""
    Else
      TxtNombre = "": TxtId.Text = ""
    End If
    CerrarRecorset rstUniversal
  Else
    TxtNombre = "": TxtId.Text = ""
  End If
End Sub

Private Function OrdenamientoGrilla() As String
  If OptGuia.Value = True Then
    OrdenamientoGrilla = "order by Guia"
  ElseIf OptFecha.Value = True Then
    OrdenamientoGrilla = "order by FhEntradaBodega"
  End If
End Function
