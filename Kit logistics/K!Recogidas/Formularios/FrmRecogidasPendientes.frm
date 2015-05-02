VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmRecogidasPendientes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Recogidas pendientes"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   13380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkEjecutar 
      Caption         =   "Ejecutar al iniciar"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton CmdVerReporte 
      Caption         =   "Ver reporte"
      Height          =   255
      Left            =   6600
      TabIndex        =   3
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton CmdActualizar 
      Caption         =   "Ver / Actualizar"
      Height          =   255
      Left            =   8640
      TabIndex        =   2
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   255
      Left            =   11160
      TabIndex        =   1
      Top             =   5520
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid GrillaRecogidas 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13095
      _ExtentX        =   23098
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "IdAnuncio"
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
         DataField       =   "FhAnuncio"
         Caption         =   "Fecha"
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
         DataField       =   "IdCliente"
         Caption         =   "Nit"
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
         DataField       =   "Anunciante"
         Caption         =   "Anunciante"
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
         DataField       =   "IdRuta"
         Caption         =   "Ruta"
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
         DataField       =   "Asignacion"
         Caption         =   "Asig"
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
         DataField       =   "FhRecogida"
         Caption         =   "FhRec"
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
      BeginProperty Column08 
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
      BeginProperty Column09 
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
      BeginProperty Column10 
         DataField       =   "Programada"
         Caption         =   "P"
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
         DataField       =   "Estado"
         Caption         =   "ES"
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
         DataField       =   "Efectiva"
         Caption         =   "EF"
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
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   3014.929
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   464.882
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column08 
            Locked          =   -1  'True
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column09 
            Locked          =   -1  'True
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column10 
            Locked          =   -1  'True
            ColumnWidth     =   269.858
         EndProperty
         BeginProperty Column11 
            Locked          =   -1  'True
            ColumnWidth     =   285.165
         EndProperty
         BeginProperty Column12 
            Locked          =   -1  'True
            ColumnWidth     =   329.953
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmRecogidasPendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstTemp As New ADODB.Recordset
Private Sub CmdActualizar_Click()
  If rstTemp.State = adStateOpen Then
    rstTemp.Close
  End If
  rstTemp.Open "SELECT IdAnuncio, FhAnuncio, IdCliente, Anunciante, IdRuta, Vehiculo, FhRecogida, Unidades, KilosReales, KilosVol, Programada, Estado, Efectiva, Coperaciones From Anuncios where Efectiva=0 and FhRecogida<'" & Date & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Set GrillaRecogidas.DataSource = rstTemp
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub CmdVerReporte_Click()
    Mostrar_Reporte CnnPrincipal, 2, "Select*from anuncios where efectiva=0 and FhRecogida<'" & Date & "'", "RECOGIDAS PENDIENTES DE FECHAS ANTERIORES", 2
End Sub

Private Sub Form_Load()
  ChkEjecutar.Value = GetSetting("Kit Logistics", "Recogidas", "Ini_Rec_Pend", 0)
  rstTemp.CursorLocation = adUseClient
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveSetting "Kit logistics", "Recogidas", "Ini_Rec_Pend", ChkEjecutar.Value
  Set rstTemp = Nothing
End Sub

