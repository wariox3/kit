VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmBuscarFacturas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filtrar facturas...."
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11160
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNroResultados 
      Height          =   285
      Left            =   10320
      TabIndex        =   22
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   9240
      TabIndex        =   20
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Filtrar"
      Height          =   255
      Left            =   7200
      TabIndex        =   19
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tener en cuenta el rango de fecha"
      Height          =   375
      Left            =   8520
      TabIndex        =   18
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Frame FraFecha 
      Caption         =   "Fecha"
      Height          =   1335
      Left            =   4560
      TabIndex        =   12
      Top             =   4680
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "Cambiar"
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   720
         TabIndex        =   16
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   720
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   465
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Frame FraFactura 
      Caption         =   "Factura"
      Height          =   1335
      Left            =   1560
      TabIndex        =   5
      Top             =   4680
      Width           =   2775
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "ID Cliente:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   465
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
      Begin VB.OptionButton OptFecha 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Rango"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton OptCliente 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7858
      _Version        =   393216
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
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Nro resultados:"
      Height          =   195
      Index           =   5
      Left            =   9120
      TabIndex        =   21
      Top             =   5160
      Width           =   1065
   End
End
Attribute VB_Name = "FrmBuscarFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSalir_Click()
  Unload Me
End Sub
