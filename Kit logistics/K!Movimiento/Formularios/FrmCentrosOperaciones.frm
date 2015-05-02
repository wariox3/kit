VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmBuscarDespachoReparto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buscar Despachos de reparto"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir resultados"
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   4320
      Width           =   2055
   End
   Begin VB.OptionButton OptEmpiece 
      Caption         =   "E&mpiece"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   3600
      Width           =   975
   End
   Begin VB.OptionButton OptExacto 
      Caption         =   "&Exacto"
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   3600
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.ComboBox CboCriterio 
      Height          =   315
      ItemData        =   "FrmCentrosOperaciones.frx":0000
      Left            =   0
      List            =   "FrmCentrosOperaciones.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3240
      Width           =   3135
   End
   Begin VB.TextBox TxtBuscamos 
      Height          =   315
      Left            =   3240
      MaxLength       =   100
      TabIndex        =   1
      Top             =   3240
      Width           =   5175
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Filtrar"
      Height          =   315
      Left            =   8520
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin VB.Frame FraFechas 
      Height          =   1215
      Left            =   0
      TabIndex        =   8
      Top             =   3840
      Width           =   3135
      Begin VB.TextBox TxtDesde 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox TxtHasta 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox ChkConFecha 
         Alignment       =   1  'Right Justify
         Caption         =   "Con fecha"
         Height          =   195
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LblDesde 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   840
         Width           =   465
      End
   End
   Begin VB.CommandButton CmdCambiarFechas 
      Caption         =   "Cambiar fechas"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton CmdPanelInfoDespacho 
      Caption         =   "Panel info Despacho"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Aceptar / Salir"
      Height          =   255
      Left            =   7560
      TabIndex        =   5
      Top             =   4800
      Width           =   1935
   End
   Begin VB.OptionButton OptContenga 
      Caption         =   "&Contenga"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox TxtNroResultados 
      Height          =   285
      Left            =   7560
      MaxLength       =   4
      TabIndex        =   3
      Top             =   3960
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid GrillaGuias 
      Height          =   3015
      Left            =   0
      TabIndex        =   14
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   5318
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "OrdDespacho"
         Caption         =   "Despacho"
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
         DataField       =   "FhExpedicion"
         Caption         =   "Fh Expe"
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
         DataField       =   "FhCumplidos"
         Caption         =   "Fh Cump"
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
      BeginProperty Column03 
         DataField       =   "IdVehiculo"
         Caption         =   "Vehiculo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "IdEncargado"
         Caption         =   "Encargado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "HH:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
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
      BeginProperty Column06 
         DataField       =   "Estado"
         Caption         =   "Est"
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
         DataField       =   "Unidades"
         Caption         =   "Unds"
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
      BeginProperty Column08 
         DataField       =   "KilosReales"
         Caption         =   "K Real"
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
      BeginProperty Column09 
         DataField       =   "ContraEntregas"
         Caption         =   "CEntregas"
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
      BeginProperty Column10 
         DataField       =   "Tipo"
         Caption         =   "T"
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
      BeginProperty Column11 
         DataField       =   "CO"
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
      BeginProperty Column12 
         DataField       =   "Observaciones"
         Caption         =   "Observaciones"
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
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column07 
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
   Begin VB.Label LblMensajes 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   3600
      Width           =   5175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Numero de resultados:"
      Height          =   195
      Left            =   5880
      TabIndex        =   15
      Top             =   3960
      Width           =   1590
   End
End
Attribute VB_Name = "FrmBuscarDespachoReparto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstTemporal As New ADODB.Recordset
Private Sub CboCriterio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub CmdBuscar_Click()
  Dim Consulta As String
  LblMensajes = "Buscando... se mostraran " & TxtNroResultados & " registros"
  If TxtBuscamos.Text <> "" Then
    Consulta = "Select Top " & Val(TxtNroResultados) & " * from DespachosReparto where "
    FufuSt = DevCampos()
    If Val(Me.Tag) = 1 Then
      If IsNumeric(TxtBuscamos) Then
        Consulta = Consulta & DevCampos() & " = " & TxtBuscamos
      Else
        MsgBox "El valor para este campo de criterio debe ser numerico, por favor verifique la informacion", vbCritical
        TxtBuscamos.SetFocus
        Exit Sub
      End If
    ElseIf Val(Me.Tag) = 2 Then
      If OptExacto.Value = True Then
        Consulta = Consulta & DevCampos & " = '" & TxtBuscamos & "'"
      Else
        If OptEmpiece.Value = True Then
          Consulta = Consulta & DevCampos & " like '" & TxtBuscamos & "%'"
        Else
          Consulta = Consulta & DevCampos & " like '%" & TxtBuscamos & "%'"
        End If
      End If
    Else
      Consulta = Consulta & DevCampos() & " >= '" & TxtDesde.Text & "' and " & DevCampos() & " <= '" & TxtHasta.Text & "'"
    End If
    
    If ChkConFecha.Value = 1 Then
      If CboCriterio.ListIndex <> 6 And CboCriterio.ListIndex <> 7 Then
        Consulta = Consulta & " and FhExpedicion >= '" & TxtDesde.Text & "' and FhExpedicion <= '" & TxtHasta.Text & "'"
      End If
    End If
    'MsgBox Consulta
    AbrirRecorset rstTemporal, Consulta, adOpenForwardOnly, adLockReadOnly
    Set GrillaGuias.DataSource = rstTemporal
    LblMensajes.Caption = "Se encontro " & rstTemporal.RecordCount & " registros"
    TxtBuscamos.SetFocus
  Else
    MsgBox "Debe digitar algun dato para buscar", vbCritical
    TxtBuscamos.SetFocus
  End If
End Sub

Private Sub CmdCambiarFechas_Click()
  If Principal.ToolEspecial1.AbrirDevFechas("Digite el rango de fechas", "Digite las nuevas fechas", 2) = True Then
    TxtDesde = Principal.ToolEspecial1.Fecha1
    TxtHasta = Principal.ToolEspecial1.Fecha2
  End If
End Sub

Private Sub CmdSalir_Click()
  Set rstTemporal = Nothing
  Unload Me
End Sub

Private Sub Form_Load()
  rstTemporal.CursorLocation = adUseClient
  TxtDesde = Date
  TxtHasta = Date
  CboCriterio.ListIndex = 0
  TxtNroResultados = SacarDatoString(17)
End Sub
Private Function DevCampos() As String
  Select Case CboCriterio.ListIndex
    Case 0  'Despacho
      DevCampos = "OrdDespacho"
      Me.Tag = "1"
    Case 1  'Fecha de expedicion
      DevCampos = "FhExpedicion"
      Me.Tag = "3"
    Case 2  'Fecha de cumplidos
      DevCampos = "FhCumplidos"
      Me.Tag = "3"
    Case 3  'Vehiculo
      DevCampos = "IdVehiculo"
      Me.Tag = "2"
    Case 4  'Encargado
      DevCampos = "IdEncargado"
      Me.Tag = "2"
    Case 5  'Ruta
      DevCampos = "IdRuta"
      Me.Tag = "1"
  End Select
End Function
Private Sub TxtBuscamos_GotFocus()
  EnfocarT TxtBuscamos
End Sub
Private Sub TxtBuscamos_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub TxtNroResultados_GotFocus()
  EnfocarT TxtNroResultados
End Sub

Private Sub TxtNroResultados_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtNroResultados, KeyAscii, 1
End Sub
