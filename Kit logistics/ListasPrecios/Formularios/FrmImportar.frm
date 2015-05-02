VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmImportar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Importar datos de listas de precios"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   12510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCargar 
      Caption         =   "Cargar"
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Frame FraLista 
      Enabled         =   0   'False
      Height          =   4935
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   12255
      Begin MSDataGridLib.DataGrid GrillaPrecios 
         Height          =   4575
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   8070
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
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
            DataField       =   "IdCiudad"
            Caption         =   "IDCiu"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "NmCiudad"
            Caption         =   "Ciudad"
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
            DataField       =   "IdProducto"
            Caption         =   "IDPro"
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
            DataField       =   "NmProducto"
            Caption         =   "Producto"
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
            DataField       =   "VrKilo"
            Caption         =   "Vr Kilo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "VrUnidad"
            Caption         =   "Vr Unidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "VrTonelada"
            Caption         =   "Vr Tonelada"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "KTope"
            Caption         =   "K Tope"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "VrKTope"
            Caption         =   "Vr K Tope"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "VrKAdicional"
            Caption         =   "Vr Adiciona"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   524.976
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1409.953
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1665.071
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraTipLista 
      Caption         =   "Tipo de lista"
      Height          =   975
      Left            =   1920
      TabIndex        =   11
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton Option2 
         Caption         =   "Existente"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nueva"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdImportar 
      Caption         =   "Importar"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Frame FraBuscar 
      Height          =   4935
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   12255
      Begin VB.FileListBox FileArchivo 
         Height          =   4185
         Left            =   6360
         Pattern         =   "*.csv"
         TabIndex        =   9
         Top             =   240
         Width           =   5775
      End
      Begin VB.DirListBox DirRuta 
         Height          =   4590
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   480
      Width           =   5415
      Begin VB.TextBox TxtNmListaprecios 
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame FraDestino 
      Caption         =   "Destino"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton OptBD 
         Caption         =   "Base de datos"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptArchivo 
         Caption         =   "Archivo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   255
      Left            =   10800
      TabIndex        =   0
      Top             =   6240
      Width           =   1575
   End
End
Attribute VB_Name = "FrmImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CnnArTexto As New ADODB.Connection
Dim rstArchivo As New ADODB.Recordset

Private Sub CmdCargar_Click()
On Error GoTo ElErr
  If FileArchivo.FileName <> "" Then
    CnnArTexto.Open "DRIVER={Microsoft Text Driver (*.txt; *.csv)}; DBQ=" & DirRuta.Path & ";", "", ""
    rstArchivo.Open "select * from " & FileArchivo.FileName, CnnArTexto, adOpenStatic, adLockReadOnly, adCmdText
    Set GrillaPrecios.DataSource = rstArchivo
    Set CnnArTexto = Nothing
    FraBuscar.Visible = False
    FraBuscar.Enabled = False
    FraLista.Enabled = True
    FraLista.Visible = True
    CmdImportar.Enabled = True
    CmdCargar.Enabled = False
  Else
    MsgBox "No hay un archivo seleccionado, debe seleccionar un archivo", vbCritical
  End If
ElErr:
  If Err.Number > 0 Then
    MsgBox Err.Number & " - " & Err.Description
    Select Case Err.Number
      Case 1
    End Select
  End If

End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub


Private Sub DirRuta_Change()
  FileArchivo.Path = DirRuta.Path
End Sub

Private Sub Form_Load()
  CnnArTexto.CursorLocation = adUseClient
  rstArchivo.CursorLocation = adUseClient
End Sub
