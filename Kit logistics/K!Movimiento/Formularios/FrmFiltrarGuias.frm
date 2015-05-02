VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBuscarGuias 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filtrar Guias..."
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   12615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExportar 
      Height          =   615
      Left            =   11880
      Picture         =   "FrmFiltrarGuias.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6480
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Entregar"
      Height          =   1095
      Left            =   4800
      TabIndex        =   26
      Top             =   6240
      Width           =   3135
      Begin MSComCtl2.DTPicker DTPHoraEntrega 
         Height          =   375
         Left            =   1560
         TabIndex        =   29
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16842754
         CurrentDate     =   39775
      End
      Begin MSComCtl2.DTPicker DTPFechaEntregada 
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16842753
         CurrentDate     =   39775
      End
      Begin VB.CommandButton CmdEntregarGuia 
         Caption         =   "Entregar guia"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   2895
      End
   End
   Begin VB.CommandButton CmdMasInformacion 
      Caption         =   "Mas informacion"
      Height          =   255
      Left            =   10680
      TabIndex        =   25
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "Guardar"
      Height          =   255
      Left            =   10680
      TabIndex        =   24
      Top             =   7440
      Width           =   1815
   End
   Begin VB.TextBox TxtOptSeleccionado 
      Height          =   285
      Left            =   11880
      TabIndex        =   23
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TxtNroReg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9840
      TabIndex        =   22
      Text            =   "50"
      Top             =   7080
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid GrillaGuias 
      Height          =   5535
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   9763
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
      ColumnCount     =   13
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
      BeginProperty Column02 
         DataField       =   "FhEntradaBodega"
         Caption         =   "Fecha"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Cliente"
         Caption         =   "Cliente"
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
      BeginProperty Column09 
         DataField       =   "IdFactura"
         Caption         =   "Factura"
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
         DataField       =   "IdDespacho"
         Caption         =   "Despacho"
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
         DataField       =   "DirDestinatario"
         Caption         =   "Direccion destinatario"
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
         DataField       =   "TelDestinatario"
         Caption         =   "Tel destinatario"
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   2520
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column08 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column09 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column10 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column11 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column12 
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   8760
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   10680
      TabIndex        =   19
      Top             =   7800
      Width           =   1815
   End
   Begin VB.TextBox TxtVerConsulta 
      Height          =   615
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   7440
      Width           =   9975
   End
   Begin VB.TextBox TxtSelFin 
      Height          =   285
      Left            =   11880
      TabIndex        =   16
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TxtSelIni 
      Height          =   285
      Left            =   11880
      TabIndex        =   15
      Text            =   "="
      Top             =   6720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Caption         =   "Criterio"
      Height          =   1095
      Left            =   3000
      TabIndex        =   10
      Top             =   6240
      Width           =   1575
      Begin VB.OptionButton OptCri 
         Caption         =   "&Igual"
         Height          =   230
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   160
         Value           =   -1  'True
         Width           =   855
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
         Caption         =   "&Contenga"
         Height          =   230
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton OptCri 
         Caption         =   "&Empiece"
         Height          =   230
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   390
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdCambiarFechas 
      Caption         =   "Cambiar"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Filtrar"
      Height          =   255
      Left            =   10680
      TabIndex        =   2
      Top             =   5760
      Width           =   1815
   End
   Begin VB.TextBox TxtBuscamos 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   5760
      Width           =   7575
   End
   Begin VB.ComboBox CboOpcion 
      Height          =   315
      ItemData        =   "FrmFiltrarGuias.frx":0374
      Left            =   120
      List            =   "FrmFiltrarGuias.frx":039C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   5760
      Width           =   2775
   End
   Begin VB.CheckBox ChkConFecha 
      Caption         =   "Por fecha"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Frame FraFechas 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   6240
      Width           =   2775
      Begin VB.TextBox TxtFh2 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TxtFh1 
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   510
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nro Registros"
      Height          =   195
      Left            =   8760
      TabIndex        =   21
      Top             =   7080
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "SQL:"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   7440
      Width           =   360
   End
End
Attribute VB_Name = "FrmBuscarGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstBuscar As New ADODB.Recordset
Dim Consulta As String
Private Sub CboOpcion_Click()
  OptCri_Click (Val(TxtOptSeleccionado.Text))
  'VerC
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
  If Err.Number <> 0 Then MsgBox "La consulta esta mal estructurada, no se va a mostrar ninguna informacion", vbCritical
End Sub
Function SacarCampo(Opcion As Byte)
  Select Case Opcion
    Case 0
      SacarCampo = "Guia "
    Case 1
      SacarCampo = "FhEntradaBodega"
    Case 2
      SacarCampo = "FhEntregaMercancia "
    Case 3
      SacarCampo = "DocCliente "
    Case 4
      SacarCampo = "NmDestinatario "
    Case 5
      SacarCampo = "NmCiudad "
    Case 6
      SacarCampo = "IdRuta "
    Case 7
      SacarCampo = "Remitente "
    Case 8
      SacarCampo = "Cliente "
    Case 9
      SacarCampo = "IdFactura "
    Case 10
      SacarCampo = "IdDespacho "
    Case 11
      SacarCampo = "Cuenta "
  End Select
End Function

Private Sub CmdCambiarFechas_Click()
  If Principal.ToolConsultas1.AbrirDevFechas("Rango de fechas", "Digite un rango de fechas para el filtro de la factura", 2) = True Then
    TxtFh1.Text = Principal.ToolConsultas1.Fecha1
    TxtFh2.Text = Principal.ToolConsultas1.Fecha2
  End If
  VerC
End Sub

Private Sub CmdEntregarGuia_Click()
If rstBuscar.State = adStateOpen Then
  If rstBuscar.EOF = False Then
    If MsgBox("¿Desea entregar la guia " & rstBuscar.Fields("Guia") & " con fecha " & Format(DTPFechaEntregada, "dd/mm/yyyy") & " y hora " & Format(DTPHoraEntrega.Value, "h:m:s"), vbQuestion + vbYesNo, "Entregar guia") = vbYes Then
        AbrirRecorset rstUniversal, "Update Guias Set Entregada=1, FhEntregaMercancia='" & Format(DTPFechaEntregada.Value, "yyyy/mm/dd") & " " & Format(DTPHoraEntrega.Value, "h:m:s") & "' where Guia=" & rstBuscar.Fields("Guia"), CnnPrincipal, adOpenDynamic, adLockOptimistic
    End If
  Else
    MsgBox "Debe seleccionar una guia", vbCritical
  End If
Else
  MsgBox "Debe buscar algo primero", vbCritical
End If
End Sub

Private Sub CmdExportar_Click()
If rstBuscar.State = adStateOpen Then
  If rstBuscar.EOF = False Then
    ExportarExcel rstBuscar
  End If
End If
End Sub

Private Sub CmdMasInformacion_Click()
If rstBuscar.State = adStateOpen Then
  If rstBuscar.EOF = False Then
    FufuLo = rstBuscar.Fields("Guia")
    FrmInfoGuia.Show 1
  Else
    MsgBox "Debe seleccionar una guia", vbCritical
  End If
Else
  MsgBox "Debe buscar algo primero", vbCritical
End If
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  rstBuscar.CursorLocation = adUseClient
  CboOpcion.ListIndex = 0
  Consulta = "Select Guias.*, Ciudades.NmCiudad from Guias, Ciudades where (Guias.IdCiuDestino=Ciudades.IdCiudad) and "
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
        Case 0, 6, 9, 10
          TxtSelIni.Text = "="
          TxtSelFin.Text = ""
        Case 3, 4, 5, 7, 8, 11
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
    If ChkConFecha.Value = 1 Then TxtVerConsulta.Text = TxtVerConsulta.Text & " and (FhEntradaBodega>='" & Format(TxtFh1, "yy-mm-dd") & " 00:00:00' and FhEntradaBodega <='" & Format(TxtFh2, "yy-mm-dd") & " 23:59:00')"
  Case 1, 2
    TxtVerConsulta.Text = Consulta & SacarCampo(CboOpcion.ListIndex) & ">='" & Format(TxtFh1, "yy-mm-dd") & " 00:00:00' and " & SacarCampo(CboOpcion.ListIndex) & " <='" & Format(TxtFh2, "yy-mm-dd") & " 23:59:00'"
  Case 3, 4, 5, 7, 8, 11
    TxtVerConsulta.Text = Consulta & SacarCampo(CboOpcion.ListIndex) & TxtSelIni & TxtBuscamos & TxtSelFin
    If ChkConFecha.Value = 1 Then TxtVerConsulta.Text = TxtVerConsulta.Text & " and (FhEntradaBodega>='" & Format(TxtFh1, "yy-mm-dd") & " 00:00:00' and FhEntradaBodega <='" & Format(TxtFh2, "yy-mm-dd") & " 23:59:00')"
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
