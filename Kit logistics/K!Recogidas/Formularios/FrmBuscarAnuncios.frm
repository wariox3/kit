VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmBuscarAnuncios 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buscar anuncios..."
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   12570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdVerReporte 
      Caption         =   "Ver Reporte"
      Height          =   255
      Left            =   10680
      TabIndex        =   24
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CheckBox ChkConFecha 
      Caption         =   "Por fecha"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   4800
      Width           =   1095
   End
   Begin VB.ComboBox CboOpcion 
      Height          =   315
      ItemData        =   "FrmBuscarAnuncios.frx":0000
      Left            =   120
      List            =   "FrmBuscarAnuncios.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox TxtBuscamos 
      Height          =   285
      Left            =   3000
      TabIndex        =   15
      Top             =   4320
      Width           =   7575
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Filtrar"
      Height          =   255
      Left            =   10680
      TabIndex        =   14
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de busqueda"
      Height          =   1095
      Left            =   4680
      TabIndex        =   11
      Top             =   4800
      Width           =   1695
      Begin VB.OptionButton OptTBRapido 
         Caption         =   "Rapido"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptTBCompleta 
         Caption         =   "Completa"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdCambiarFechas 
      Caption         =   "Cambiar"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   4800
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "Criterio"
      Height          =   1095
      Left            =   3000
      TabIndex        =   5
      Top             =   4800
      Width           =   1575
      Begin VB.OptionButton OptCri 
         Caption         =   "Empiece"
         Height          =   230
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   390
         Width           =   1095
      End
      Begin VB.OptionButton OptCri 
         Caption         =   "Contenga"
         Height          =   230
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton OptCri 
         Caption         =   "Termine"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   810
         Width           =   1215
      End
      Begin VB.OptionButton OptCri 
         Caption         =   "Igual"
         Height          =   230
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   160
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox TxtSelIni 
      Height          =   285
      Left            =   7800
      TabIndex        =   4
      Text            =   "="
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TxtSelFin 
      Height          =   285
      Left            =   7800
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TxtVerConsulta 
      Height          =   615
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   6000
      Width           =   9975
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   10680
      TabIndex        =   1
      Top             =   6360
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid GrillaAnuncios 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   7223
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
      ColumnCount     =   19
      BeginProperty Column00 
         DataField       =   "IdAnuncio"
         Caption         =   "Anuncio"
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
         DataField       =   "FhAnuncio"
         Caption         =   "Fecha"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "IdCliente"
         Caption         =   "ID Cliente"
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
         DataField       =   "Anunciante"
         Caption         =   "Anunciante"
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
         DataField       =   "DirAnunciante"
         Caption         =   "Dir anunciante"
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
         DataField       =   "TelAnunciante"
         Caption         =   "Telefono"
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
      BeginProperty Column06 
         DataField       =   "IdRuta"
         Caption         =   "Ruta"
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
         DataField       =   "Vehiculo"
         Caption         =   "Asig"
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
         DataField       =   "FhRecogida"
         Caption         =   "Hr Rec"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "FhRecogida"
         Caption         =   "Fh Rec"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Unidades"
         Caption         =   "Und"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "KilosReales"
         Caption         =   "KR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "KilosVol"
         Caption         =   "KV"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "Comentarios"
         Caption         =   "Comentarios"
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
      BeginProperty Column14 
         DataField       =   "Programada"
         Caption         =   "P"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   4
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "Estado"
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
      BeginProperty Column16 
         DataField       =   "Efectiva"
         Caption         =   "E"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   4
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "TiempoEfectiva"
         Caption         =   "Tiempo efectiva"
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
      BeginProperty Column18 
         DataField       =   "Coperaciones"
         Caption         =   "CO"
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
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1964.976
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   1934.929
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   464.882
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column08 
            Locked          =   -1  'True
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column09 
            Locked          =   -1  'True
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column10 
            Locked          =   -1  'True
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column11 
            Locked          =   -1  'True
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column12 
            Locked          =   -1  'True
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column13 
            Locked          =   -1  'True
            ColumnWidth     =   1920.189
         EndProperty
         BeginProperty Column14 
            Locked          =   -1  'True
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column15 
            Locked          =   -1  'True
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column16 
            Locked          =   -1  'True
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column17 
            Locked          =   -1  'True
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column18 
            Locked          =   -1  'True
            ColumnWidth     =   315.213
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraFechas 
      Height          =   1095
      Left            =   120
      TabIndex        =   18
      Top             =   4800
      Width           =   2775
      Begin VB.TextBox TxtFh1 
         Height          =   285
         Left            =   720
         TabIndex        =   20
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox TxtFh2 
         Height          =   285
         Left            =   720
         TabIndex        =   19
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "SQL:"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   6000
      Width           =   360
   End
End
Attribute VB_Name = "FrmBuscarAnuncios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstBuscar As New ADODB.Recordset
Dim Consulta As String
Private Sub CboOpcion_Click()
  OptCri_Click (0)
End Sub
Private Sub ChkConFecha_Click()
  VerC
End Sub
Private Sub CmdBuscar_Click()
On Error GoTo SQLMalo
  If rstBuscar.State = adStateOpen Then rstBuscar.Close
  If TxtBuscamos.Text <> "" Or CboOpcion.ListIndex = 1 Or CboOpcion.ListIndex = 2 Then
    rstBuscar.Open TxtVerConsulta.Text, CnnPrincipal, adOpenDynamic, adLockOptimistic
    Set GrillaAnuncios.DataSource = rstBuscar
    If rstBuscar.RecordCount > 0 Then MsgBox rstBuscar.RecordCount & " registros encontrados", vbInformation
  Else
    MsgBox "Debe digitar algo para buscar", vbCritical
  End If
SQLMalo:
  If Err.Number <> 0 Then MsgBox "La consulta esta mal estructurada, no se va a mostrar ninguna informacion", vbCritical
End Sub
Function SacarCampo(Opcion As Byte)
  Select Case Opcion
    Case 0
      SacarCampo = "IdAnuncio "
    Case 1
      SacarCampo = "IdCliente "
    Case 2
      SacarCampo = "IdRuta "
    Case 3
      SacarCampo = "Vehiculo "
    Case 4
      SacarCampo = "Programada "
    Case 5
      SacarCampo = "Efectiva "
    Case 6
      SacarCampo = "COperaciones "
  End Select
End Function

Private Sub CmdCambiarFechas_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para el filtro de la factura", 2) = True Then
    TxtFh1.Text = Principal.ToolConsultas1.Fecha1
    TxtFh2.Text = Principal.ToolConsultas1.Fecha2
  End If
  VerC
End Sub
Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub CmdVerReporte_Click()
  If rstBuscar.State = adStateOpen Then
    If rstBuscar.RecordCount > 0 Then
      Mostrar_Reporte CnnPrincipal, 5, TxtVerConsulta.Text, "LISTADO ANUNCIOS Y/O RECOGIDAS", 2
    End If
  End If
End Sub

Private Sub Form_Load()
  rstBuscar.CursorLocation = adUseClient
  CboOpcion.ListIndex = 0
  Consulta = "Select*from Anuncios where "
  TxtFh1.Text = Date
  TxtFh2.Text = Date
  VerC
End Sub

Private Sub OptCri_Click(Index As Integer)
  IniFin Index
End Sub
Private Sub IniFin(Tipo As Integer)
  Select Case Tipo
    Case 0
      Select Case CboOpcion.ListIndex
        Case 0, 2, 3, 4, 5, 6
          TxtSelIni.Text = "="
          TxtSelFin.Text = ""
        Case 1
          TxtSelIni.Text = "='"
          TxtSelFin.Text = "'"
      End Select
    Case 1
      TxtSelIni.Text = "Like '%"
      TxtSelFin.Text = "'"
    Case 2
      TxtSelIni.Text = "Like '%"
      TxtSelFin.Text = "%'"
    Case 3
      TxtSelIni.Text = "Like '"
      TxtSelFin.Text = "%'"
  End Select
  VerC
End Sub
Private Sub VerC()
Select Case CboOpcion.ListIndex
  Case 0, 2, 3, 4, 5, 6
    TxtVerConsulta.Text = Consulta & SacarCampo(CboOpcion.ListIndex) & TxtSelIni & TxtBuscamos & TxtSelFin
    If ChkConFecha.Value = 1 Then TxtVerConsulta.Text = TxtVerConsulta.Text & " and (FhFactura>=#" & Format(TxtFh1, "mm/dd/yyyy") & "# and FhFactura<=#" & Format(TxtFh2, "mm/dd/yyyy") & "#)"
  'Case 1, 2
  '  TxtVerConsulta.Text = Consulta & SacarCampo(CboOpcion.ListIndex) & ">=#" & Format(TxtFh1, "mm/dd/yyyy") & "# and " & SacarCampo(CboOpcion.ListIndex) & "<=#" & Format(TxtFh2, "mm/dd/yyyy") & "#"
  Case 1
    TxtVerConsulta.Text = Consulta & SacarCampo(CboOpcion.ListIndex) & TxtSelIni & TxtBuscamos & TxtSelFin
    If ChkConFecha.Value = 1 Then TxtVerConsulta.Text = TxtVerConsulta.Text & " and (FhFactura>=#" & Format(TxtFh1, "mm/dd/yyyy") & "# and FhFactura<=#" & Format(TxtFh2, "mm/dd/yyyy") & "#)"
  End Select
End Sub

Private Sub OptTBCompleta_Click()
  Consulta = "SELECT Facturas.*, Propietarios.NmPropietario FROM Facturas LEFT JOIN Propietarios ON Facturas.IdPropietario = Propietarios.IdPropietario where "
  VerC
End Sub

Private Sub OptTBRapido_Click()
  Consulta = "Select*from Guias where "
  VerC
End Sub

Private Sub TxtBuscamos_Change()
  VerC
End Sub
Private Sub TxtBuscamos_KeyPress(KeyAscii As Integer)
  VerC
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
