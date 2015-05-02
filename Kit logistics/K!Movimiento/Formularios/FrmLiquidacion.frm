VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmLiquidacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Liquidacion de la remision..."
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptAdicional 
      Caption         =   "A&dicional"
      Enabled         =   0   'False
      Height          =   195
      Left            =   5400
      TabIndex        =   86
      Top             =   4080
      Width           =   975
   End
   Begin VB.OptionButton OptKilo 
      Caption         =   "&Kilo"
      Enabled         =   0   'False
      Height          =   270
      Left            =   3720
      TabIndex        =   85
      Top             =   4080
      Width           =   690
   End
   Begin VB.OptionButton OptUnidad 
      Caption         =   "&Unidad"
      Enabled         =   0   'False
      Height          =   270
      Left            =   4440
      TabIndex        =   84
      Top             =   4080
      Width           =   930
   End
   Begin MSDataGridLib.DataGrid GrillaLista 
      Height          =   975
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1720
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "NmProducto"
         Caption         =   "Producto"
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
         DataField       =   "VrKilo"
         Caption         =   "Kilo"
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
      BeginProperty Column02 
         DataField       =   "VrUnidad"
         Caption         =   "Unidad"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "KTope"
         Caption         =   "Tope"
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
      BeginProperty Column04 
         DataField       =   "VrKTope"
         Caption         =   "Vr Tope"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "VrKAdicional"
         Caption         =   "Adicional"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#.##0"
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
            ColumnWidth     =   2220.094
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1065.26
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraTot 
      Enabled         =   0   'False
      Height          =   615
      Left            =   7080
      TabIndex        =   76
      Top             =   5000
      Width           =   4575
      Begin MSMask.MaskEdBox TxtVrManejo 
         Height          =   255
         Left            =   960
         TabIndex        =   79
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0;($#,##0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtVrFlete 
         Height          =   255
         Left            =   3120
         TabIndex        =   80
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0;($#,##0)"
         PromptChar      =   "_"
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Vr Manejo:"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Width           =   765
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Vr Flete:"
         Height          =   195
         Index           =   13
         Left            =   2520
         TabIndex        =   77
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.CommandButton CmdAjuste 
      Caption         =   "Ajuste"
      Height          =   375
      Left            =   6720
      TabIndex        =   75
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox TxtMinDespacho 
      Height          =   285
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   74
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CheckBox ChkPermitirListaGeneral 
      Caption         =   "Permitir lista general"
      Height          =   255
      Left            =   120
      TabIndex        =   72
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox TxtPorManejo 
      Height          =   285
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   65
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox TxtMinUnidad 
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10080
      TabIndex        =   13
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8280
      TabIndex        =   58
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "Agre&gar"
      Height          =   255
      Left            =   8280
      TabIndex        =   11
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton CmdRetirar 
      Caption         =   "&Retirar"
      Height          =   255
      Left            =   9960
      TabIndex        =   47
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Frame FrProductos 
      Height          =   1455
      Left            =   120
      TabIndex        =   23
      Top             =   1080
      Width           =   11535
      Begin VB.TextBox TxtLote 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   15
         TabIndex        =   1
         Top             =   270
         Width           =   3735
      End
      Begin VB.TextBox TxtIdProducto 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   3
         TabIndex        =   2
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox TxtIdEmpaque 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   3
         TabIndex        =   3
         Top             =   990
         Width           =   495
      End
      Begin MSMask.MaskEdBox TxtPrecios 
         Height          =   255
         Index           =   0
         Left            =   5400
         TabIndex        =   24
         Top             =   165
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtCantidad 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Left            =   10320
         TabIndex        =   4
         Top             =   180
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   5
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPrecios 
         Height          =   255
         Index           =   1
         Left            =   5400
         TabIndex        =   25
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPrecios 
         Height          =   255
         Index           =   2
         Left            =   5400
         TabIndex        =   26
         Top             =   1125
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtAlto 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   10
         Top             =   1125
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   5
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtLargo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   9
         Top             =   810
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   5
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtAncho 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   8
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   5
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtKFac 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Left            =   10320
         TabIndex        =   7
         Top             =   1125
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   5
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtKR 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Left            =   10320
         TabIndex        =   6
         Top             =   810
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   5
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtKVol 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Left            =   10320
         TabIndex        =   5
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   16777215
         MaxLength       =   5
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPrecios 
         Height          =   255
         Index           =   4
         Left            =   6000
         TabIndex        =   68
         Top             =   165
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPrecios 
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   70
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPrecios 
         Height          =   255
         Index           =   5
         Left            =   7800
         TabIndex        =   87
         Top             =   165
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         Format          =   "0"
         PromptChar      =   "_"
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Min:"
         Height          =   195
         Index           =   1
         Left            =   7440
         TabIndex        =   88
         Top             =   165
         Width           =   300
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Unidad:"
         Height          =   195
         Index           =   0
         Left            =   4860
         TabIndex        =   69
         Top             =   840
         Width           =   555
      End
      Begin VB.Label LblConsulta 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   45
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label LblConsulta 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   44
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   37
         Left            =   4950
         TabIndex        =   43
         Top             =   165
         Width           =   465
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "$"
         Height          =   195
         Index           =   41
         Left            =   5880
         TabIndex        =   42
         Top             =   165
         Width           =   90
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         Height          =   195
         Index           =   35
         Left            =   120
         TabIndex        =   41
         Top             =   630
         Width           =   645
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   195
         Index           =   34
         Left            =   360
         TabIndex        =   40
         Top             =   270
         Width           =   315
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Empaque"
         Height          =   195
         Index           =   36
         Left            =   120
         TabIndex        =   39
         Top             =   990
         Width           =   675
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Index           =   48
         Left            =   9555
         TabIndex        =   38
         Top             =   150
         Width           =   630
      End
      Begin VB.Label LblUniversal 
         Caption         =   "Kilo:"
         Height          =   255
         Index           =   39
         Left            =   5040
         TabIndex        =   37
         Top             =   1125
         Width           =   375
      End
      Begin VB.Label LblUniversal 
         Caption         =   "Adicional:"
         Height          =   255
         Index           =   38
         Left            =   4680
         TabIndex        =   36
         Top             =   480
         Width           =   735
      End
      Begin VB.Label LblUniversal 
         Caption         =   "Cms"
         Height          =   255
         Index           =   47
         Left            =   8760
         TabIndex        =   35
         Top             =   1125
         Width           =   375
      End
      Begin VB.Label LblUniversal 
         Caption         =   "Cms"
         Height          =   255
         Index           =   46
         Left            =   8760
         TabIndex        =   34
         Top             =   810
         Width           =   375
      End
      Begin VB.Label LblUniversal 
         Caption         =   "Cms"
         Height          =   255
         Index           =   45
         Left            =   8760
         TabIndex        =   33
         Top             =   480
         Width           =   375
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Alto"
         Height          =   195
         Index           =   44
         Left            =   7440
         TabIndex        =   32
         Top             =   1125
         Width           =   270
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Largo"
         Height          =   195
         Index           =   43
         Left            =   7305
         TabIndex        =   31
         Top             =   810
         Width           =   405
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Ancho"
         Height          =   195
         Index           =   42
         Left            =   7245
         TabIndex        =   30
         Top             =   480
         Width           =   465
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "K Facturados"
         Height          =   195
         Index           =   51
         Left            =   9240
         TabIndex        =   29
         Top             =   1125
         Width           =   945
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "K real"
         Height          =   195
         Index           =   50
         Left            =   9780
         TabIndex        =   28
         Top             =   810
         Width           =   405
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "K volumen"
         Height          =   195
         Index           =   49
         Left            =   9435
         TabIndex        =   27
         Top             =   495
         Width           =   750
      End
   End
   Begin VB.Frame FraListas 
      Height          =   1035
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   3945
      Begin VB.TextBox TxtNmListaGeneral 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   82
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox TxtIdListaPreciosGeneral 
         Height          =   285
         Left            =   555
         TabIndex        =   81
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox TxtNmListaPrecios 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   71
         Top             =   160
         Width           =   2655
      End
      Begin VB.TextBox TxtIdListaPrecioC 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   560
         MaxLength       =   200
         TabIndex        =   67
         Top             =   160
         Width           =   495
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "Gal:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   83
         Top             =   480
         Width           =   285
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "Lista:"
         Height          =   195
         Index           =   41
         Left            =   120
         TabIndex        =   18
         Top             =   165
         Width           =   375
      End
   End
   Begin VB.Frame FraDesManejo 
      Enabled         =   0   'False
      Height          =   520
      Left            =   120
      TabIndex        =   19
      Top             =   5640
      Width           =   4935
      Begin VB.TextBox TxtDctoUni 
         Height          =   285
         Left            =   2280
         TabIndex        =   64
         Top             =   160
         Width           =   735
      End
      Begin VB.TextBox TxtMinimos 
         Height          =   285
         Left            =   3960
         TabIndex        =   60
         Top             =   160
         Width           =   735
      End
      Begin VB.TextBox TxtDctoKil 
         Height          =   285
         Left            =   840
         TabIndex        =   21
         Top             =   160
         Width           =   735
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Dcto Uni:"
         Height          =   195
         Index           =   3
         Left            =   1560
         TabIndex        =   63
         Top             =   165
         Width           =   675
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Minimos:"
         Height          =   195
         Index           =   2
         Left            =   3240
         TabIndex        =   59
         Top             =   165
         Width           =   615
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Dcto Kil:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   165
         Width           =   600
      End
   End
   Begin MSComctlLib.ListView LstTem 
      Height          =   1335
      Left            =   120
      TabIndex        =   46
      Top             =   2640
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Lote"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Producto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ID"
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Empaque"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Ancho"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Largo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Alto"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Cant"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "K Real"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "K Vol"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Kilos Fac"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Vr Flete"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSMask.MaskEdBox TxtDeclarado 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   4560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   15
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Frame FraLiquidacion 
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   48
      Top             =   5000
      Width           =   6855
      Begin MSMask.MaskEdBox TxtKFacturar 
         Height          =   255
         Left            =   6000
         TabIndex        =   49
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtKVolumen 
         Height          =   255
         Left            =   2280
         TabIndex        =   50
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtKReales 
         Height          =   255
         Left            =   3960
         TabIndex        =   51
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtUnidades 
         Height          =   255
         Left            =   840
         TabIndex        =   52
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Unidades:"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   720
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "K Reales:"
         Height          =   195
         Index           =   10
         Left            =   3240
         TabIndex        =   55
         Top             =   240
         Width           =   690
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "K Vol:"
         Height          =   195
         Index           =   9
         Left            =   1800
         TabIndex        =   54
         Top             =   240
         Width           =   420
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "K Facturados:"
         Height          =   195
         Index           =   8
         Left            =   4920
         TabIndex        =   53
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.Label LblCiudadOrigen 
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
      Height          =   255
      Left            =   5640
      TabIndex        =   90
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Origen:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   4920
      TabIndex        =   89
      Top             =   6360
      Width           =   630
   End
   Begin VB.Line Separador 
      BorderWidth     =   2
      Index           =   2
      X1              =   11640
      X2              =   120
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Label Label1 
      Caption         =   "Min X Despacho:"
      Height          =   195
      Left            =   8880
      TabIndex        =   73
      Top             =   4560
      Width           =   1230
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "% Manejo:"
      Height          =   195
      Index           =   1
      Left            =   3120
      TabIndex        =   66
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label LblTituloMinimo 
      AutoSize        =   -1  'True
      Caption         =   "Min X unidad:"
      Height          =   195
      Left            =   5760
      TabIndex        =   61
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Declarado:"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   57
      Top             =   4560
      Width           =   780
   End
   Begin VB.Line Separador 
      BorderWidth     =   2
      Index           =   1
      X1              =   11640
      X2              =   120
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Separador 
      BorderWidth     =   2
      Index           =   0
      X1              =   11640
      X2              =   120
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label LblCliente 
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
      Height          =   255
      Left            =   960
      TabIndex        =   22
      Top             =   6360
      Width           =   3735
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   6360
      Width           =   660
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Destino:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   8280
      TabIndex        =   15
      Top             =   6360
      Width           =   720
   End
   Begin VB.Label LblCiudadDestino 
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
      Height          =   255
      Left            =   9120
      TabIndex        =   14
      Top             =   6360
      Width           =   2535
   End
End
Attribute VB_Name = "FrmLiquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ajuste As Boolean
Dim FleteViejo As Double
Dim ManejoViejo As Double
Dim rstListaPrecios As New ADODB.Recordset

Private Sub CmdAceptar_Click()
FrmRemisiones.Campo(15) = Val(TxtUnidades)
FrmRemisiones.Campo(16) = Val(TxtKReales)
FrmRemisiones.Campo(18) = Val(TxtKVolumen)
FrmRemisiones.Campo(17) = Val(TxtKFacturar)
FrmRemisiones.Campo(12) = Val(TxtDeclarado)
FrmRemisiones.Campo(14) = Val(TxtVrManejo)
FrmRemisiones.Campo(13) = Val(TxtVrFlete)
Erase MProductos
  For II = 1 To LstTem.ListItems.Count
    MProductos(II).Lote = LstTem.ListItems.Item(II)
    MProductos(II).IdProducto = LstTem.ListItems.Item(II).SubItems(1)
    MProductos(II).NmProducto = LstTem.ListItems.Item(II).SubItems(2)
    MProductos(II).IdEmpaque = LstTem.ListItems.Item(II).SubItems(3)
    MProductos(II).NmEmpaque = LstTem.ListItems.Item(II).SubItems(4)
    MProductos(II).Ancho = Val(LstTem.ListItems.Item(II).SubItems(5))
    MProductos(II).Largo = Val(LstTem.ListItems.Item(II).SubItems(6))
    MProductos(II).Alto = Val(LstTem.ListItems.Item(II).SubItems(7))
    MProductos(II).Cantidad = LstTem.ListItems.Item(II).SubItems(8)
    MProductos(II).KilosReales = Val(LstTem.ListItems.Item(II).SubItems(9))
    MProductos(II).KilosVol = Val(LstTem.ListItems.Item(II).SubItems(10))
    MProductos(II).KilosFacturados = Val(LstTem.ListItems.Item(II).SubItems(11))
    MProductos(II).VrFlete = LstTem.ListItems.Item(II).SubItems(12)
  Next
IdClienteViejo = FrmRemisiones.Campo(3).Text
Unload Me
End Sub

Private Sub CmdAgregar_Click()
If Validacion2 = False Then Exit Sub
Set Item = LstTem.ListItems.Add(, , TxtLote.Text)
    Item.SubItems(1) = TxtIdProducto
    Item.SubItems(2) = LblConsulta(1)
    Item.SubItems(3) = TxtIdEmpaque
    Item.SubItems(4) = LblConsulta(2)
    Item.SubItems(5) = TxtAncho.Text
    Item.SubItems(6) = TxtLargo.Text
    Item.SubItems(7) = TxtAlto.Text
    Item.SubItems(8) = TxtCantidad.Text
    Item.SubItems(9) = TxtKR.Text
    Item.SubItems(10) = TxtKVol.Text
    Item.SubItems(11) = TxtKFac.Text

If OptKilo.value = True Then
    If Val(TxtKFac.Text) < Val(TxtMinimos.Text) Then If MsgBox("Los kilos minimos son: " & TxtMinimos & ", Y el producto esta siendo facturado por:" & TxtKFac & " ¿DESEA CONTINUAR CON ESTOS KILOS?" & Chr(13) & "- SI ingresarlo con estos kilos" & Chr(13) & "- NO cambiarle los kilos facturados", vbYesNo + vbInformation, "¿Valor menor que el mínimo?") = vbNo Then TxtCantidad.SetFocus: Exit Sub
    TxtVrFlete = Val(TxtVrFlete.Text) + (Val(TxtKFac) * Val(TxtPrecios(2))) * (1 - ((Val(TxtDctoKil) / 100)))
    Item.SubItems(12) = (Val(TxtKFac) * Val(TxtPrecios(2))) * (1 - ((Val(TxtDctoKil) / 100)))
    OptUnidad.Enabled = False: OptAdicional.Enabled = False
End If

If OptUnidad.value = True Then
  TxtVrFlete = Val(TxtVrFlete.Text) + (Val(TxtCantidad.Text) * Val(TxtPrecios(3).Text) * (1 - ((ValNum(TxtDctoUni) / 100))))
  Item.SubItems(12) = (Val(TxtCantidad.Text) * Val(TxtPrecios(3).Text)) * (1 - ((ValNum(TxtDctoUni) / 100)))
  OptKilo.Enabled = False: OptAdicional.Enabled = False
End If

If OptAdicional.value = True Then
  If Val(TxtKFac.Text) / Val(TxtCantidad) > Val(TxtPrecios(0).Text) Then
    TxtVrFlete = Val(TxtVrFlete.Text) + ((Val(TxtCantidad) * Val(TxtPrecios(4))) + ((TxtKFac - (TxtPrecios(0) * TxtCantidad)) * TxtPrecios(1)))
    
    Item.SubItems(12) = ((Val(TxtCantidad) * Val(TxtPrecios(4))) + ((TxtKFac - (TxtPrecios(0) * TxtCantidad)) * TxtPrecios(1)))
    'Item.SubItems(12) = ((Val(TxtCantidad) * Val(TxtPrecios(4))) + ((TxtKFac - (TxtPrecios(0) * TxtCantidad)) * TxtPrecios(4)))
  Else
    TxtVrFlete = Val(TxtVrFlete.Text) + (Val(TxtPrecios(4).Text) * Val(TxtCantidad))
    Item.SubItems(12) = Val(TxtPrecios(4).Text) * Val(TxtCantidad)
  End If
OptKilo.Enabled = False: OptUnidad.Enabled = False
End If

TxtKReales = Val(TxtKReales) + Val(TxtKR.Text)
TxtKFacturar = Val(TxtKFacturar.Text) + Val(TxtKFac.Text)
TxtKVolumen = Val(TxtKVolumen.Text) + Val(TxtKVol.Text)
TxtUnidades = Val(TxtUnidades.Text) + Val(TxtCantidad.Text)
limpiar
TxtLote.SetFocus
End Sub

Private Sub CmdAjuste_Click()
  Ajuste = True
  FrProductos.Enabled = False
  TxtDeclarado.Enabled = False
  FleteViejo = Val(TxtVrFlete.Text)
  ManejoViejo = Val(TxtVrManejo.Text)
  FraTot.Enabled = True
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub CmdRetirar_Click()
II = 1
While II <= LstTem.ListItems.Count
  If LstTem.ListItems(II).Checked = True Then
    TxtUnidades = ValNum(TxtUnidades) - LstTem.ListItems(II).SubItems(8)
    TxtKReales = ValNum(TxtKReales) - LstTem.ListItems(II).SubItems(9)
    TxtKFacturar = ValNum(TxtKFacturar) - LstTem.ListItems(II).SubItems(11)
    TxtKVolumen = ValNum(TxtKVolumen) - LstTem.ListItems(II).SubItems(10)
    TxtVrFlete = ValNum(TxtVrFlete) - LstTem.ListItems(II).SubItems(12)
    LstTem.ListItems.Remove (II)
  Else
    II = II + 1
  End If
Wend
End Sub

Private Sub Form_Load()
  Dim strSql As String
  rstListaPrecios.CursorLocation = adUseClient
  AbrirRecorset rstUniversal, "Select ListaPreciosGeneral from configuracion", CnnPrincipal, adOpenDynamic, adLockOptimistic
    TxtIdListaPreciosGeneral.Text = rstUniversal.Fields("ListaPreciosGeneral")
  CerrarRecorset rstUniversal
  AbrirRecorset rstUniversal, "Select listasprecios.* from listasprecios where IdListaPrecios=" & Val(TxtIdListaPreciosGeneral.Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstUniversal.RecordCount > 0 Then
      TxtNmListaGeneral.Text = rstUniversal.Fields("NmListaPrecios")
    End If
  CerrarRecorset rstUniversal
  
  LblCliente = FrmRemisiones.LblConsulta(0)
  LblCiudadOrigen = FrmRemisiones.LblConsulta(3)
  LblCiudadDestino = FrmRemisiones.LblConsulta(1)
  
  AbrirRecorset rstUniversal, "SELECT Id, DctoK, DctoU, Minimos, MinUniManejo, MinDesManejo, ManKilo, ManUni, ManAdicional, PorManejo, ListaPrecios, PermiteListaGral, NoAplicarDctoReexpediciones From Negociaciones where Id=" & Val(FrmRemisiones.Campo(3)), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = True Then
    AbrirRecorset rstUniversal, "SELECT PorManejo, Dcto, IdListaPrecios, Minimos, MinUniManejo From ParametrizacionLiquidaciones", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    ActivarOptions True, True, True
  Else
    ActivarOptions rstUniversal!ManKilo, rstUniversal!ManUni, rstUniversal!ManAdicional
  End If
    TxtDctoKil = rstUniversal!DctoK
    If DevCheck(rstUniversal!NoAplicarDctoReexpediciones) = 1 Then
      Dim rstCiudades As New ADODB.Recordset
      rstCiudades.CursorLocation = adUseClient
      AbrirRecorset rstCiudades, "SELECT Reexpedicion FROM ciudades WHERE Reexpedicion = 1 AND IdCiudad = " & Val(FrmRemisiones.Campo(8).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstCiudades.RecordCount > 0 Then
          TxtDctoKil.Text = 0
      End If
    End If
    TxtDctoUni = rstUniversal!DctoU
    TxtIdListaPrecioC = rstUniversal!ListaPrecios
    TxtPorManejo = rstUniversal!PorManejo
    TxtMinimos = rstUniversal!Minimos
    TxtMinUnidad = rstUniversal!MinUniManejo
    TxtMinDespacho = rstUniversal!MinDesManejo
    ChkPermitirListaGeneral.value = DevCheck(rstUniversal!PermiteListaGral)
  CerrarRecorset rstUniversal
  If Val(TxtIdListaPrecioC.Text) <> 0 Then TxtNmListaPrecios.Text = DevResBus("SELECT IdListaPrecios, NmListaPrecios From ListasPrecios where IdListaPrecios=" & Val(TxtIdListaPrecioC), "NmListaPrecios", CnnPrincipal)
  If Val(TxtIdListaPreciosGeneral.Text) <> 0 Then TxtNmListaGeneral.Text = DevResBus("SELECT IdListaPrecios, NmListaPrecios From ListasPrecios where IdListaPrecios=" & Val(TxtIdListaPreciosGeneral), "NmListaPrecios", CnnPrincipal)
  Ajuste = False
  strSql = "SELECT listaspreciosciudades.*, NmProducto " & _
           "FROM listaspreciosciudades " & _
           "LEFT JOIN productos ON listaspreciosciudades.IdProducto = productos.IdProducto " & _
           "WHERE IdListaPrecios = " & TxtIdListaPrecioC.Text & " AND IdCiudadOrigen = " & FrmRemisiones.Campo(30).Text & " AND IdCiudad = " & FrmRemisiones.Campo(8).Text
  AbrirRecorset rstListaPrecios, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
  Set GrillaLista.DataSource = rstListaPrecios
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstListaPrecios = Nothing
End Sub

Private Sub GrillaLista_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub GrillaLista_LostFocus()
  If rstListaPrecios.State = adStateOpen Then
    If rstListaPrecios.EOF = False Then
      TxtIdProducto.Text = rstListaPrecios.Fields("IdProducto")
      TxtIdProducto.SetFocus
    End If
  End If
End Sub

Private Sub OptAdicional_KeyPress(KeyAscii As Integer)
  TxtCantidad.SetFocus
End Sub

Private Sub OptKilo_KeyPress(KeyAscii As Integer)
  TxtCantidad.SetFocus
End Sub

Private Sub OptUnidad_KeyPress(KeyAscii As Integer)
  TxtCantidad.SetFocus
End Sub

Private Sub TxtAlto_Change()
  CalculoVolumen
End Sub

Private Sub TxtAlto_GotFocus()
  EnfocarM TxtAlto
End Sub

Private Sub TxtAlto_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtAlto, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub TxtAlto_Validate(Cancel As Boolean)
  CalculoVolumen
End Sub
Private Sub TxtAncho_Change()
  CalculoVolumen
End Sub

Private Sub TxtAncho_GotFocus()
  EnfocarM TxtAncho
End Sub

Private Sub TxtAncho_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtAncho, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub TxtAncho_Validate(Cancel As Boolean)
  CalculoVolumen
End Sub

Private Sub TxtCantidad_GotFocus()
  EnfocarM TxtCantidad
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtCantidad, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtCantidad_LostFocus()
  CalculoKilosFacturar
End Sub
Private Sub TxtDeclarado_GotFocus()
  EnfocarM TxtDeclarado
End Sub

Private Sub TxtDeclarado_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtDeclarado, KeyAscii, 1
  If KeyAscii = 13 Then CmdAceptar.SetFocus
End Sub

Private Sub TxtDeclarado_LostFocus()
  Dim MjoDcldo As Currency, MjoUnidad As Currency
  MjoDcldo = (ValNum(TxtDeclarado) * ValNum(TxtPorManejo)) / 100
  MjoUnidad = ValNum(TxtUnidades) * ValNum(TxtMinUnidad)
  If MjoDcldo > MjoUnidad Then
    TxtVrManejo = Int(MjoDcldo + 0.5)
  Else
    TxtVrManejo = MjoUnidad
  End If
  If Val(TxtMinDespacho) > MjoDcldo And Val(TxtMinDespacho) > MjoUnidad Then
    TxtVrManejo = Val(TxtMinDespacho)
  End If
End Sub
Private Sub TxtIdEmpaque_GotFocus()
  EnfocarT TxtIdEmpaque
End Sub

Private Sub TxtIdEmpaque_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
      Principal.ToolConsultas1.AbrirConsultaGral "IdEmpaque", "NmEmpaque", "Empaques", CnnPrincipal
      TxtIdEmpaque.Text = Principal.ToolConsultas1.DatLo
  End If
End Sub
Private Sub TxtIdEmpaque_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtIdEmpaque, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub TxtIdEmpaque_Validate(Cancel As Boolean)
  AbrirRecorset rstUniversal, "SELECT IdEmpaque, NmEmpaque FROM Empaques where IdEmpaque=" & Val(TxtIdEmpaque), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = False Then
    LblConsulta(2) = rstUniversal!NmEmpaque & ""
  Else
    LblConsulta(2) = "": TxtIdEmpaque = ""
  End If
  CerrarRecorset rstUniversal
End Sub

Private Sub TxtIdListaPrecioC_GotFocus()
  EnfocarT TxtIdListaPrecioC
End Sub

Private Sub TxtIdListaPrecioC_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsulta 11, CnnPrincipal
    TxtIdListaPrecioC.Text = Principal.ToolConsultas1.DatLo
  End If
End Sub

Private Sub TxtIdListaPrecioC_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtIdListaPrecioC, KeyAscii, 1
  If KeyAscii = 13 Then TxtLote.SetFocus
End Sub

Private Sub TxtIdListaPrecioC_LostFocus()
  If Val(TxtIdListaPrecioC.Text) <> 0 Then
    AbrirRecorset rstUniversal, "SELECT IdListaPrecios, NmListaPrecios From ListasPrecios where IdListaPrecios=" & Val(TxtIdListaPrecioC.Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      TxtNmListaPrecios.Text = rstUniversal!NmListaPrecios & ""
    Else
      TxtNmListaPrecios.Text = "": TxtIdListaPrecioC.Text = ""
    End If
    CerrarRecorset rstUniversal
  Else
    TxtNmListaPrecios.Text = "": TxtIdListaPrecioC.Text = ""
  End If
End Sub

Private Sub TxtIdProducto_GotFocus()
  EnfocarT TxtIdProducto
End Sub

Private Sub TxtIdProducto_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
      Principal.ToolConsultas1.AbrirConsultaGral "IdProducto", "NmProducto", "Productos", CnnPrincipal
      TxtIdProducto.Text = Principal.ToolConsultas1.DatLo
  End If
End Sub
Private Sub TxtIdProducto_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtIdProducto, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub TxtIdProducto_Validate(Cancel As Boolean)
  If Val(TxtIdListaPrecioC.Text) <> 0 Then
    If Val(TxtIdProducto) <> 0 Then
      AbrirRecorset rstUniversal, "SELECT IdProducto, NmProducto FROM Productos where IdProducto=" & Val(TxtIdProducto), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        LblConsulta(1) = rstUniversal!NmProducto & ""
        Dim TemListaPrecios As TipListaPrecios
        If ChkPermitirListaGeneral.value = 0 Then
          TemListaPrecios = DevListaPrecios(Val(FrmRemisiones.Campo(8)), Val(TxtIdProducto.Text), Val(TxtIdListaPrecioC.Text), Val(TxtIdListaPreciosGeneral.Text), False)
        Else
          TemListaPrecios = DevListaPrecios(Val(FrmRemisiones.Campo(8)), Val(TxtIdProducto.Text), Val(TxtIdListaPrecioC.Text), Val(TxtIdListaPreciosGeneral.Text), True)
        End If
        
        If TemListaPrecios.Devuelve = True Then
          TxtPrecios(0) = TemListaPrecios.KTope
          TxtPrecios(1) = TemListaPrecios.VrKdicional
          TxtPrecios(2) = TemListaPrecios.VrKilo
          TxtPrecios(3) = TemListaPrecios.VrUnidad
          TxtPrecios(4) = TemListaPrecios.VrKTope
          TxtPrecios(5) = TemListaPrecios.Minimos
        End If
      Else
        LblConsulta(1) = "": TxtIdProducto = ""
      End If
      CerrarRecorset rstUniversal
    Else
      LblConsulta(1) = "": TxtIdProducto.Text = ""
    End If
  Else
    MsgBox "Este cliente no tiene lista de precios, debe asignarle una lista de precios para poder liquidar el flete", vbCritical
  End If
End Sub

Private Sub TxtKFac_GotFocus()
  EnfocarM TxtKFac
End Sub

Private Sub TxtKFac_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtKFac, KeyAscii, 1
  If KeyAscii = 13 Then CmdAgregar.SetFocus
End Sub

Private Sub TxtKR_GotFocus()
  EnfocarM TxtKR
End Sub

Private Sub TxtKR_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtKR, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtKR_LostFocus()
  CalculoKilosFacturar
End Sub

Private Sub TxtKVol_GotFocus()
  EnfocarM TxtKVol
End Sub

Private Sub TxtKVol_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtKVol, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub TxtKVol_LostFocus()
  CalculoKilosFacturar
End Sub

Private Sub TxtLargo_Change()
  CalculoVolumen
End Sub

Private Sub TxtLargo_GotFocus()
  EnfocarM TxtLargo
End Sub

Private Sub TxtLargo_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtLargo, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtLargo_Validate(Cancel As Boolean)
  CalculoVolumen
End Sub

Private Sub TxtLote_GotFocus()
  EnfocarT TxtLote
End Sub

Private Sub TxtLote_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then TxtDeclarado.SetFocus
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Function Validacion2() As Boolean
  If Val(TxtIdListaPrecioC.Text) <> 0 Then
    If Val(TxtIdProducto.Text) <> 0 Then
      If Val(TxtIdEmpaque) <> 0 Then
        If Val(TxtKR.Text) <> 0 Then
          If Val(TxtKVol.Text) <> 0 Then
            If Val(TxtKFac.Text) <> 0 Then
              If Val(TxtCantidad.Text) <> 0 Then
                If OptKilo.value = True Then
                  If Val(TxtPrecios(2).Text) <> 0 Then
                    Validacion2 = True
                  Else
                    MsgBox "No hay un precio de kilo para liquidar el producto, mire las listas de precios": TxtIdProducto.SetFocus: Validacion2 = False: Exit Function
                  End If
                End If
              
                If OptUnidad.value = True Then
                  If Val(TxtPrecios(3).Text) <> 0 Then
                    Validacion2 = True
                  Else
                    MsgBox "No hay un precio de unidad para liquidar el producto, mire las listas de precios": TxtIdProducto.SetFocus:  Validacion2 = False: Exit Function
                  End If
                End If
              
                If OptAdicional.value = True Then
                  If Val(TxtPrecios(4).Text) <> 0 And Val(TxtPrecios(0).Text) <> 0 Then
                    Validacion2 = True
                  Else
                    MsgBox "No hay Kilos Tope o valor kilos Tope para operar, compruebe las listas de precios": TxtIdProducto.SetFocus:  Validacion2 = False: Exit Function
                  End If
                End If
              Else
                MsgBox "Debe digitar una cantidad de producto": TxtCantidad.SetFocus: Validacion2 = False:  Exit Function
              End If
            Else
              MsgBox "No a definido cuantos kilos va a facturar de este producto":  TxtKFac.SetFocus:  Validacion2 = False:  Exit Function
            End If
          Else
            MsgBox "No a definido cuantos kilos volumen tiene este producto":  TxtKVol.SetFocus:  Validacion2 = False:  Exit Function
          End If
        Else
          MsgBox "No a definido cuantos kilos reales tiene este producto":  TxtKR.SetFocus:  Validacion2 = False:  Exit Function
        End If
      Else
        MsgBox "Especifique un empaque para el producto": TxtIdEmpaque.SetFocus: Validacion2 = False: Exit Function
      End If
    Else
      MsgBox "No ha especificado un producto": TxtIdProducto.SetFocus:  Validacion2 = False:  Exit Function
    End If
  Else
    MsgBox "No hay una lista de precios para liquidar": TxtIdListaPrecioC.SetFocus:  Validacion2 = False:  Exit Function
  End If
End Function
Sub ActivarOptions(MKilo As Boolean, MUnidad As Boolean, MAdicional As Boolean)
      If MKilo = True And MUnidad = True And MAdicional = True Then
        OptUnidad.value = False: OptAdicional.value = False:  OptKilo.value = True: OptUnidad.Enabled = True: OptAdicional.Enabled = True:   OptKilo.Enabled = True
      ElseIf (MKilo = True And MUnidad = True And MAdicional = False) Then
        OptUnidad.value = False: OptAdicional.value = False:  OptKilo.value = True: OptUnidad.Enabled = True: OptAdicional.Enabled = False:   OptKilo.Enabled = True
      ElseIf (MKilo = True And MUnidad = False And MAdicional = True) Then
        OptUnidad.value = False: OptAdicional.value = False:  OptKilo.value = True: OptUnidad.Enabled = False: OptAdicional.Enabled = True:   OptKilo.Enabled = True
      ElseIf (MKilo = False And MUnidad = True And MAdicional = True) Then
        OptUnidad.value = True: OptAdicional.value = False:  OptKilo.value = False: OptUnidad.Enabled = True: OptAdicional.Enabled = True:   OptKilo.Enabled = False
      ElseIf (MKilo = True And MUnidad = False And MAdicional = False) Then
        OptUnidad.value = False: OptAdicional.value = False:  OptKilo.value = True: OptUnidad.Enabled = False: OptAdicional.Enabled = False:   OptKilo.Enabled = True
      ElseIf (MKilo = False And MUnidad = True And MAdicional = False) Then
        OptUnidad.value = True: OptAdicional.value = False: OptKilo.value = False: OptUnidad.Enabled = True: OptAdicional.Enabled = False:   OptKilo.Enabled = False
      ElseIf (MKilo = False And MUnidad = False And MAdicional = True) Then
        OptUnidad.value = False: OptAdicional.value = True: OptKilo.value = False: OptUnidad.Enabled = False: OptAdicional.Enabled = True:   OptKilo.Enabled = False
      ElseIf MKilo = False And MUnidad = False And MAdicional = False Then
        OptUnidad.value = False: OptAdicional.value = False:  OptKilo.value = True: OptUnidad.Enabled = True: OptAdicional.Enabled = True:   OptKilo.Enabled = True
        MsgBox "Este cliente no maneja por ninguna opcion"
      End If
End Sub

Private Sub limpiar()
  TxtLote.Text = ""
  TxtIdProducto.Text = ""
  TxtIdEmpaque.Text = ""
  For II = 0 To 5
    TxtPrecios(II).Text = ""
  Next
  TxtAncho.Text = ""
  TxtLargo.Text = ""
  TxtAlto.Text = ""
  TxtCantidad.Text = ""
  TxtKVol.Text = ""
  TxtKR.Text = ""
  TxtKFac.Text = ""
  LblConsulta(1).Caption = ""
  LblConsulta(2).Caption = ""
End Sub
Sub CalculoKilosFacturar()
If OptKilo.value = True Then
  If Val(TxtKVol.Text) > Val(TxtKFac.Text) Then TxtKFac = TxtKVol
  If Val(TxtKR.Text) > Val(TxtKFac.Text) Then TxtKFac.Text = TxtKR
  If Val(TxtPrecios(5).Text) > 0 Then
    If (Val(TxtPrecios(5).Text) * Val(TxtCantidad.Text)) > Val(TxtKFac.Text) Then TxtKFac = (Val(TxtPrecios(5).Text) * Val(TxtCantidad.Text))
  Else
    If (Val(TxtMinimos) * Val(TxtCantidad.Text)) > Val(TxtKFac.Text) Then TxtKFac = (Val(TxtMinimos) * Val(TxtCantidad.Text))
  End If
End If
If OptAdicional.value = True Then
  TxtKFac = TxtPrecios(0) * Val(TxtCantidad)
  If Val(TxtKVol.Text) > Val(TxtKFac.Text) Then TxtKFac = TxtKVol
  If Val(TxtKR.Text) > Val(TxtKFac.Text) Then TxtKFac.Text = TxtKR
End If
If OptUnidad.value = True Then
  If Val(TxtKVol.Text) > Val(TxtKFac.Text) Then TxtKFac = TxtKVol
  If Val(TxtKR.Text) > Val(TxtKFac.Text) Then TxtKFac.Text = TxtKR
End If
End Sub
Sub CalculoVolumen()
  TxtKVol = ((Val(TxtAncho.Text) / 100) * (Val(TxtLargo.Text) / 100) * (Val(TxtAlto.Text) / 100) * 400)
End Sub

