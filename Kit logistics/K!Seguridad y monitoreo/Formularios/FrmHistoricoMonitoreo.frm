VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmHistoricoMonitoreo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Historico de monitoreo"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   255
      Left            =   6000
      TabIndex        =   8
      Top             =   4440
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DGMonitoreos 
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5106
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "ID"
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
         DataField       =   "Orden"
         Caption         =   "Ord"
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
         DataField       =   "Tipo"
         Caption         =   "Tipo"
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
         DataField       =   "Estado"
         Caption         =   "Est"
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
         DataField       =   "FhHrSalida"
         Caption         =   "Salida"
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
         DataField       =   "Vehiculo"
         Caption         =   "Vehiculo"
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
         DataField       =   "Destino"
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
      BeginProperty Column07 
         DataField       =   "UltReporte"
         Caption         =   "Ult Rep"
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
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   2069.858
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            ColumnWidth     =   2145.26
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraCriterios 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9135
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   7920
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPHasta 
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Format          =   16908289
         CurrentDate     =   40282
      End
      Begin MSComCtl2.DTPicker DTPDesde 
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Format          =   16908289
         CurrentDate     =   40282
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   465
      End
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   7680
      TabIndex        =   0
      Top             =   4440
      Width           =   1575
   End
End
Attribute VB_Name = "FrmHistoricoMonitoreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim consulta As String
Dim rstBuscar As New ADODB.Recordset

Private Sub CmdBuscar_Click()
If rstBuscar.State = adStateOpen Then rstBuscar.Close
rstBuscar.Open consulta & " where FhHrSalida>='" & Format(DTPDesde.Value, "yy-mm-dd") & " 00:00:00' and FhHrSalida<='" & Format(DTPHasta.Value, "yy-mm-dd") & " 23:59:59'", CnnPrincipal, adOpenDynamic, adLockOptimistic
Set DGMonitoreos.DataSource = rstBuscar
End Sub

Private Sub CmdImprimir_Click()
If rstBuscar.State = adStateOpen Then
  If rstBuscar.EOF = False Then
    FufuLo = rstBuscar.Fields("ID")
    Mostrar_Reporte CnnPrincipal, 21, "Select*from sql_ism_monitoreos where IDMonitoreo=" & FufuLo, "", 2
  Else
    MsgBox "Debe seleccionar un monitoreo", vbCritical
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
  consulta = "select * from monitoreovehiculos"
  DTPDesde.Value = Date
  DTPHasta.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rstBuscar = Nothing
End Sub
