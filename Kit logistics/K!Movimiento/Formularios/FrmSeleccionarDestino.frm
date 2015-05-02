VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmSeleccionarDestino 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccionar destino..."
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdVerRutas 
      Caption         =   "Ver Rutas"
      Height          =   255
      Left            =   6840
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdVer 
      Caption         =   "Ver"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   5760
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid GrillaTramos 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Ciudad"
         Caption         =   "ID"
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
         DataField       =   "NOMBRE"
         Caption         =   "Ciudad"
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
      BeginProperty Column02 
         DataField       =   "KILOMETROSTRAMO"
         Caption         =   "Km"
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
         DataField       =   "KILOMETROSTOTAL"
         Caption         =   "Km T"
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
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   4559.811
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   675.213
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   6480
      TabIndex        =   1
      Top             =   5760
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid GrillaDestinos 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "codigoRuta"
         Caption         =   "Cod"
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
         DataField       =   "nombre"
         Caption         =   "Nombre"
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
      BeginProperty Column02 
         DataField       =   "kilometros"
         Caption         =   "Km"
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
         DataField       =   "valorxTonelada"
         Caption         =   "Vr Ton"
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
         DataField       =   "codigoDestino"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   3614.74
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   1154.835
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Destino:"
      Height          =   195
      Left            =   2760
      TabIndex        =   8
      Top             =   120
      Width           =   585
   End
   Begin VB.Label LblDestino 
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Origen:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   510
   End
   Begin VB.Label LblOrigen 
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FrmSeleccionarDestino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstRutasCaracteristicas As New ADODB.Recordset
Dim rstTramosRutas As New ADODB.Recordset

Private Sub CmdAceptar_Click()
  Dim rstCiudades As New ADODB.Recordset
  rstCiudades.CursorLocation = adUseClient
  If rstRutasCaracteristicas.State = adStateOpen And rstTramosRutas.State = adStateOpen Then
    If rstRutasCaracteristicas.EOF = False And rstTramosRutas.EOF = False Then
      AbrirRecorset rstCiudades, "Select * from ciudades where CodMinTrans='" & rstTramosRutas.Fields("CIUDAD") & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstCiudades.RecordCount > 0 Then
        FrmManifiestos.TxtCampos(35).Text = rstRutasCaracteristicas.Fields("CODIGORUTA")
        FrmManifiestos.TxtCampos(36).Text = rstRutasCaracteristicas.Fields("kilometros")
        FrmManifiestos.TxtCampos(37).Text = (rstRutasCaracteristicas.Fields("valorxTonelada") / rstRutasCaracteristicas.Fields("kilometros")) * rstTramosRutas.Fields("KILOMETROSTOTAL")
        FrmManifiestos.TxtCampos(38).Text = LblOrigen.Caption
        FrmManifiestos.TxtCampos(39).Text = rstTramosRutas.Fields("CIUDAD")
        FrmManifiestos.TxtCampos(40).Text = 0
        FrmManifiestos.TxtCampos(41).Text = rstTramosRutas.Fields("KILOMETROSTOTAL")
        'FrmManifiestos.TxtCampos(37).Text = Format((rstRutasCaracteristicas.Fields("valorxTonelada") / rstRutasCaracteristicas.Fields("CODIGORUTA")) * rstTramosRutas.Fields("KILOMETROSTOTAL"), "#,##0.00;(#,##0.00)")
        
        FufuLo = rstCiudades.Fields("IdCiudad")
        FufuLo = 1
        Unload Me
      Else
        MsgBox "No se enuentra creada la ciudad con el codigo del ministerio " & rstRutasCaracteristicas.Fields("codigoDestino") & " de la ruta que desea asignar", vbCritical
      End If
      CerrarRecorset rstCiudades
    End If
  End If
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
  FufuLo = 0
End Sub

Private Sub CmdVer_Click()
  If rstRutasCaracteristicas.EOF = False Then
    AbrirRecorset rstTramosRutas, "Select CIUDAD, NOMBRE, KILOMETROSTRAMO, KILOMETROSTOTAL from tramos_ruta where CODIGORUTA=" & rstRutasCaracteristicas.Fields("codigoRuta") & " Order By SECUENCIAOD", CnnPrincipal, adOpenDynamic, adLockOptimistic
    Set GrillaTramos.DataSource = rstTramosRutas
  End If
End Sub

Private Sub CmdVerRutas_Click()
  LlenarGrilla
End Sub

Private Sub Form_Load()
  rstRutasCaracteristicas.CursorLocation = adUseClient
  rstTramosRutas.CursorLocation = adUseClient
  LblOrigen.Caption = FufuSt
  LlenarGrilla
End Sub

Private Sub LlenarGrilla()
AbrirRecorset rstRutasCaracteristicas, "Select codigoRuta, nombre, kilometros, valorxTonelada, codigoDestino from rutas_referencia where codigoOrigen='" & LblOrigen.Caption & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
Set GrillaDestinos.DataSource = rstRutasCaracteristicas
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstRutasCaracteristicas = Nothing
  Set rstTramosRutas = Nothing
End Sub
