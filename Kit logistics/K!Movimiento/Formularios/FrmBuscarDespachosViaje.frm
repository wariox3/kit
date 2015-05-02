VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmBuscarDespachosViaje 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buscar despachos"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExportar 
      Height          =   615
      Left            =   10680
      Picture         =   "FrmBuscarDespachosViaje.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6120
      Width           =   615
   End
   Begin VB.CheckBox ChkConFecha 
      Caption         =   "Por fecha"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton CmdCambiarFechas 
      Caption         =   "Cambiar"
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   6240
      Width           =   855
   End
   Begin VB.Frame FraFechas 
      Height          =   1095
      Left            =   120
      TabIndex        =   18
      Top             =   6240
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
   Begin VB.ComboBox CboOpcion 
      Height          =   315
      ItemData        =   "FrmBuscarDespachosViaje.frx":0374
      Left            =   120
      List            =   "FrmBuscarDespachosViaje.frx":0390
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   5760
      Width           =   2775
   End
   Begin VB.TextBox TxtBuscamos 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   5760
      Width           =   7575
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Filtrar"
      Height          =   255
      Left            =   10680
      TabIndex        =   2
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Criterio"
      Height          =   1095
      Left            =   3000
      TabIndex        =   11
      Top             =   6240
      Width           =   1575
      Begin VB.OptionButton OptCri 
         Caption         =   "&Empiece"
         Height          =   230
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   390
         Width           =   1095
      End
      Begin VB.OptionButton OptCri 
         Caption         =   "&Contenga"
         Height          =   230
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton OptCri 
         Caption         =   "&Termine"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   810
         Width           =   1215
      End
      Begin VB.OptionButton OptCri 
         Caption         =   "&Igual"
         Height          =   230
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   160
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox TxtSelIni 
      Height          =   285
      Left            =   11880
      TabIndex        =   10
      Text            =   "="
      Top             =   6720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TxtSelFin 
      Height          =   285
      Left            =   11880
      TabIndex        =   9
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TxtVerConsulta 
      Height          =   615
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   7440
      Width           =   9975
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   10680
      TabIndex        =   7
      Top             =   7800
      Width           =   1815
   End
   Begin VB.TextBox TxtNroReg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9840
      TabIndex        =   5
      Text            =   "50"
      Top             =   7080
      Width           =   735
   End
   Begin VB.TextBox TxtOptSeleccionado 
      Height          =   285
      Left            =   11880
      TabIndex        =   4
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "Guardar"
      Height          =   255
      Left            =   10680
      TabIndex        =   3
      Top             =   7440
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid GrillaGuias 
      Height          =   5535
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   9763
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "SQL:"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   7440
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nro Registros"
      Height          =   195
      Left            =   8760
      TabIndex        =   23
      Top             =   7080
      Width           =   960
   End
End
Attribute VB_Name = "FrmBuscarDespachosViaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstBuscar As New ADODB.Recordset
Dim Consulta As String
Private Sub CboOpcion_Click()
  OptCri_Click (Val(TxtOptSeleccionado.Text))
End Sub
Private Sub CboOpcion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub ChkConFecha_Click()
  VerC
End Sub
Private Sub CmdBuscar_Click()
On Error GoTo SQLMalo
  If rstBuscar.State = adStateOpen Then rstBuscar.Close
  If TxtBuscamos.Text <> "" Or CboOpcion.ListIndex = 1 Or CboOpcion.ListIndex = 2 Then
    rstBuscar.Open TxtVerConsulta.Text, CnnPrincipal, adOpenDynamic, adLockOptimistic
    Set GrillaGuias.DataSource = rstBuscar
    If rstBuscar.RecordCount > 0 Then MsgBox rstBuscar.RecordCount & " registros encontrados", vbInformation
  Else
    MsgBox "Debe digitar algo para buscar", vbCritical
  End If
SQLMalo:
  If Err.Number <> 0 Then MsgBox "La consulta esta mal estructurada, no se va a mostrar ninguna informacion" & Chr(13) & "Error:" & Chr(13) & Err.Description, vbCritical
End Sub
Function SacarCampo(Opcion As Byte)
  Select Case Opcion
    Case 0
      SacarCampo = "OrdDespacho "
    Case 1
      SacarCampo = "FhExpedicion "
    Case 2
      SacarCampo = "FhCumplidos "
    Case 3
      SacarCampo = "IdManifiesto "
    Case 4
      SacarCampo = "IdVehiculo "
    Case 5
      SacarCampo = "NmConductor "
    Case 6
      SacarCampo = "IdCiudadOrigen "
    Case 7
      SacarCampo = "NmCiudad "
  End Select
End Function

Private Sub CmdCambiarFechas_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para el filtro de la factura", 2) = True Then
    TxtFh1.Text = Principal.ToolConsultas1.Fecha1
    TxtFh2.Text = Principal.ToolConsultas1.Fecha2
  End If
  VerC
End Sub


Private Sub CmdExportar_Click()
If rstBuscar.State = adStateOpen Then
  If rstBuscar.EOF = False Then
    ExportarExcel rstBuscar
  End If
End If
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  rstBuscar.CursorLocation = adUseClient
  CboOpcion.ListIndex = 0
  Consulta = "Select despachos.*, Ciudades.NmCiudad from despachos, Ciudades where (despachos.IdCiudadDestino=Ciudades.IdCiudad) and "
  TxtFh1.Text = Date
  TxtFh2.Text = Date
  VerC
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstBuscar = Nothing
End Sub

Private Sub OptCri_Click(Index As Integer)
  IniFin Index
End Sub
Private Sub IniFin(Tipo As Integer)
  Select Case Tipo
    Case 0
      Select Case CboOpcion.ListIndex
        Case 0, 3, 6
          TxtSelIni.Text = "="
          TxtSelFin.Text = ""
        Case 4, 5, 7
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
  TxtOptSeleccionado.Text = Tipo
  VerC
End Sub
Private Sub VerC()
Select Case CboOpcion.ListIndex
  Case 0, 6, 9, 10
    TxtVerConsulta.Text = Consulta & SacarCampo(CboOpcion.ListIndex) & TxtSelIni & TxtBuscamos & TxtSelFin
    If ChkConFecha.Value = 1 Then TxtVerConsulta.Text = TxtVerConsulta.Text & " and (FhExpedicion>='" & Format(TxtFh1, "yy-mm-dd") & " 00:00:00' and FhExpedicion <='" & Format(TxtFh2, "yy-mm-dd") & " 23:59:00')"
  Case 1, 2
    TxtVerConsulta.Text = Consulta & SacarCampo(CboOpcion.ListIndex) & ">='" & Format(TxtFh1, "yy-mm-dd") & " 00:00:00' and " & SacarCampo(CboOpcion.ListIndex) & " <='" & Format(TxtFh2, "yy-mm-dd") & " 23:59:00'"
  Case 3, 4, 5, 7, 8, 11
    TxtVerConsulta.Text = Consulta & SacarCampo(CboOpcion.ListIndex) & TxtSelIni & TxtBuscamos & TxtSelFin
    If ChkConFecha.Value = 1 Then TxtVerConsulta.Text = TxtVerConsulta.Text & " and (FhExpedicion>='" & Format(TxtFh1, "yy-mm-dd") & " 00:00:00' and FhExpedicion <='" & Format(TxtFh2, "yy-mm-dd") & " 23:59:00')"
  End Select
  TxtVerConsulta.Text = TxtVerConsulta.Text & " Limit " & Val(TxtNroReg.Text)
End Sub

Private Sub OptCri_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then CboOpcion.SetFocus
End Sub

Private Sub TxtBuscamos_Change()
  VerC
End Sub
Private Sub TxtBuscamos_KeyPress(KeyAscii As Integer)
  VerC
  If KeyAscii = 13 Then SendKeys vbTab
End Sub


Private Sub TxtNroReg_Change()
  VerC
End Sub

