VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmControlGuiasServicio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de guias por servicio..."
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkEntregada 
      Caption         =   "Entregada"
      Height          =   255
      Left            =   8160
      TabIndex        =   9
      Top             =   240
      Width           =   1335
   End
   Begin VB.CheckBox ChkDescargada 
      Caption         =   "Descargada"
      Height          =   255
      Left            =   8160
      TabIndex        =   8
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton CmdExportar 
      Height          =   615
      Left            =   120
      Picture         =   "FrmControlGuiasServicio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6000
      Width           =   615
   End
   Begin VB.CommandButton CmdVerDetalle 
      Caption         =   "Ver Detalle"
      Height          =   255
      Left            =   9360
      TabIndex        =   6
      Top             =   6000
      Width           =   2295
   End
   Begin VB.ComboBox CboTipoDespacho 
      Height          =   315
      ItemData        =   "FrmControlGuiasServicio.frx":0374
      Left            =   4320
      List            =   "FrmControlGuiasServicio.frx":0384
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "Consultar"
      Height          =   315
      Left            =   10080
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox CboTpServicio 
      Height          =   315
      ItemData        =   "FrmControlGuiasServicio.frx":03C0
      Left            =   1020
      List            =   "FrmControlGuiasServicio.frx":03D9
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid GrillaGuias 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9340
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "Guia"
         Caption         =   "Guia"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "FhEntradaBodega"
         Caption         =   "Fecha"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Cliente"
         Caption         =   "Cliente"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
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
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Unidades"
         Caption         =   "Und"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "OrdDespacho"
         Caption         =   "Desp"
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
         DataField       =   "FhExpedicion"
         Caption         =   "Fh Despacho"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "NmConductor"
         Caption         =   "Conductor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "IdVehiculo"
         Caption         =   "Placa"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Entregada"
         Caption         =   "E"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1980.284
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   345.26
         EndProperty
      EndProperty
   End
   Begin VB.Label LblTitulos 
      AutoSize        =   -1  'True
      Caption         =   "Tp Despacho:"
      Height          =   195
      Index           =   0
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   1020
   End
   Begin VB.Label LblTitulos 
      AutoSize        =   -1  'True
      Caption         =   "Tp Servicio:"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FrmControlGuiasServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstGuias As New ADODB.Recordset

Private Sub CmdConsultar_Click()
  LlenarGrilla
End Sub



Private Sub CmdExportar_Click()
If rstGuias.State = adStateOpen Then
  If rstGuias.EOF = False Then
    ExportarExcel rstGuias
  End If
End If
End Sub

Private Sub CmdVerDetalle_Click()
If rstGuias.State = adStateOpen Then
  If rstGuias.EOF = False Then
    FufuLo = rstGuias.Fields("Guia")
    FrmInfoGuia.Show 1
  Else
    MsgBox "Debe seleccionar una guia", vbCritical
  End If
Else
  MsgBox "Debe buscar algo primero", vbCritical
End If
End Sub

Private Sub Form_Load()
  rstGuias.CursorLocation = adUseClient
  
End Sub

Private Sub LlenarGrilla()
Dim strCriterio As String
If CboTpServicio.ListIndex <> -1 Then
  strCriterio = strCriterio & " AND TpServicio = " & CboTpServicio.ListIndex
End If

If CboTipoDespacho.ListIndex <> -1 Then
  strCriterio = strCriterio & " AND despachos.Tipo = " & CboTipoDespacho.ListIndex + 1
End If
    
  AbrirRecorset rstGuias, "SELECT Guia, FhEntradaBodega, Cliente, NmCiudad, NmDestinatario, guias.Unidades, NmConductor, FhExpedicion, Entregada, IdVehiculo, OrdDespacho, EnNovedad " & _
                          "FROM guias " & _
                          "LEFT JOIN ciudades on guias.IdCiuDestino = ciudades.IdCiudad " & _
                          "LEFT JOIN despachos on guias.IdDespacho = despachos.OrdDespacho " & _
                          "WHERE Anulada = 0 AND Descargada = " & ChkDescargada.value & " AND Entregada = " & ChkEntregada.value & _
                          strCriterio & " ORDER BY FhEntradaBodega", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Set GrillaGuias.DataSource = rstGuias
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstGuias = Nothing
End Sub
