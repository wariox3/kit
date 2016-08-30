VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmRemisiones 
   Caption         =   "Guias..."
   ClientHeight    =   11085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16680
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11085
   ScaleWidth      =   16680
   WindowState     =   2  'Maximized
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   255
      Left            =   12600
      TabIndex        =   112
      Top             =   4800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "#.##0. ;[Rojo](#.##0.);- ;"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton CmdBuscarPorDocumento 
      Caption         =   "Buscar por documento"
      Height          =   255
      Left            =   12600
      TabIndex        =   111
      Top             =   8760
      Width           =   1935
   End
   Begin VB.CommandButton CmdGenerarRecibo 
      Caption         =   "Generar recibo"
      Height          =   255
      Left            =   10440
      TabIndex        =   109
      Top             =   8400
      Width           =   1935
   End
   Begin VB.Frame FraGuiaTipo 
      Height          =   480
      Left            =   9600
      TabIndex        =   103
      Top             =   1210
      Width           =   4935
      Begin VB.Label LblGuiaTipo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   104
         Top             =   120
         Width           =   4665
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Despachos"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   12240
      TabIndex        =   98
      Top             =   5160
      Width           =   2295
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   31
         Left            =   1080
         TabIndex        =   100
         Tag             =   "1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   1065
         TabIndex        =   99
         Tag             =   "1"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Relacion:"
         Height          =   195
         Left            =   240
         TabIndex        =   102
         Top             =   240
         Width           =   675
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Despacho:"
         Height          =   195
         Index           =   17
         Left            =   120
         TabIndex        =   101
         Top             =   600
         Width           =   780
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Factura"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7920
      TabIndex        =   89
      Top             =   8280
      Width           =   2415
      Begin VB.CheckBox ChkGuiFac 
         Caption         =   "Guia factura"
         Height          =   255
         Left            =   1080
         TabIndex        =   106
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox ChkFacturada 
         Caption         =   "Corriente"
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Height          =   480
      Left            =   12240
      TabIndex        =   86
      Top             =   3120
      Width           =   2295
      Begin VB.Label LblNmUsuario 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   105
         Top             =   160
         Width           =   2025
      End
   End
   Begin VB.CommandButton CmdCambiarCO 
      Caption         =   "Cambiar Centro de Operaciones"
      Enabled         =   0   'False
      Height          =   255
      Left            =   11400
      TabIndex        =   77
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      Caption         =   "Control"
      Enabled         =   0   'False
      Height          =   615
      Left            =   480
      TabIndex        =   72
      Top             =   8280
      Width           =   7335
      Begin VB.CheckBox ChkNovedad 
         Caption         =   "Novedad"
         Height          =   255
         Left            =   6000
         TabIndex        =   83
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox ChkRelacionada 
         Caption         =   "Relacionada"
         Height          =   255
         Left            =   4800
         TabIndex        =   82
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox ChkEntregada 
         Caption         =   "Entregada"
         Height          =   255
         Left            =   1440
         TabIndex        =   76
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox ChkDescargada 
         Caption         =   "Descargada"
         Height          =   255
         Left            =   2520
         TabIndex        =   75
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox ChkDespachada 
         Caption         =   "Despachada"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox ChkAnulada 
         Caption         =   "Anulada"
         Height          =   255
         Left            =   3840
         TabIndex        =   73
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdVerProductos 
      Caption         =   "Ver Productos"
      Height          =   255
      Left            =   12600
      TabIndex        =   0
      Top             =   8400
      Width           =   1935
   End
   Begin VB.Frame FraLiquidacion 
      Enabled         =   0   'False
      Height          =   3135
      Left            =   9600
      TabIndex        =   62
      Top             =   3120
      Width           =   2535
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   26
         Left            =   1080
         TabIndex        =   23
         Tag             =   "1"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   1080
         TabIndex        =   18
         Tag             =   "1"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   1080
         TabIndex        =   19
         Tag             =   "1"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   1080
         TabIndex        =   16
         Tag             =   "1"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   1080
         TabIndex        =   17
         Tag             =   "1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   1080
         TabIndex        =   20
         Tag             =   "1"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   1080
         TabIndex        =   21
         Tag             =   "1"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   1080
         TabIndex        =   22
         Tag             =   "1"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Recaudo:"
         Height          =   195
         Index           =   28
         Left            =   240
         TabIndex        =   97
         Top             =   2760
         Width           =   705
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "K. Facturar:"
         Height          =   195
         Index           =   21
         Left            =   135
         TabIndex        =   94
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "K. Volumen:"
         Height          =   195
         Index           =   22
         Left            =   105
         TabIndex        =   93
         Top             =   960
         Width           =   855
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Kilos Reales:"
         Height          =   195
         Index           =   23
         Left            =   45
         TabIndex        =   92
         Top             =   600
         Width           =   915
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Unidades:"
         Height          =   195
         Index           =   24
         Left            =   240
         TabIndex        =   91
         Top             =   240
         Width           =   720
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Flete:"
         Height          =   195
         Index           =   16
         Left            =   570
         TabIndex        =   65
         Top             =   2040
         Width           =   390
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Declarado:"
         Height          =   195
         Index           =   20
         Left            =   180
         TabIndex        =   64
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Manejo:"
         Height          =   195
         Index           =   25
         Left            =   390
         TabIndex        =   63
         Top             =   2400
         Width           =   570
      End
   End
   Begin VB.Frame FraCO 
      Caption         =   "Centros de operacion"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   9600
      TabIndex        =   51
      Top             =   1800
      Width           =   4935
      Begin VB.TextBox TxtCOCar 
         Height          =   285
         Left            =   1320
         TabIndex        =   57
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox TxtCOIng 
         Height          =   285
         Left            =   1320
         TabIndex        =   56
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   53
         Tag             =   "1"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   720
         TabIndex        =   52
         Tag             =   "1"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "C.O Car:"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   55
         Top             =   720
         Width           =   600
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "C.O Ing:"
         Height          =   195
         Index           =   30
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Facturas"
      Enabled         =   0   'False
      Height          =   975
      Left            =   12240
      TabIndex        =   48
      Top             =   3720
      Width           =   2295
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   35
         Left            =   1080
         TabIndex        =   107
         Tag             =   "1"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   34
         Left            =   120
         TabIndex        =   88
         Tag             =   "1"
         Top             =   480
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   33
         Left            =   120
         TabIndex        =   87
         Tag             =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1080
         TabIndex        =   49
         Tag             =   "1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Abonos:"
         Height          =   195
         Index           =   14
         Left            =   360
         TabIndex        =   108
         Top             =   600
         Width           =   585
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Corriente:"
         Height          =   195
         Index           =   18
         Left            =   285
         TabIndex        =   50
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame FraDatosLectura 
      Enabled         =   0   'False
      Height          =   615
      Left            =   480
      TabIndex        =   38
      Top             =   600
      Width           =   14055
      Begin VB.TextBox Campo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/mm/yy hh:mm am/pm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   32
         Left            =   8640
         TabIndex        =   84
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Campo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/mm/yy hh:mm am/pm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   6240
         TabIndex        =   60
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Campo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/mm/yy hh:mm am/pm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   3480
         TabIndex        =   58
         Top             =   240
         Width           =   1680
      End
      Begin VB.TextBox TxtEstado 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   11880
         TabIndex        =   43
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox Campo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   42
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Campo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   11160
         TabIndex        =   41
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Fh Des:"
         Height          =   195
         Index           =   29
         Left            =   8040
         TabIndex        =   85
         Top             =   240
         Width           =   555
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Fh Entrega:"
         Height          =   195
         Index           =   26
         Left            =   5400
         TabIndex        =   61
         Top             =   240
         Width           =   825
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Fh Entrada:"
         Height          =   195
         Index           =   27
         Left            =   2520
         TabIndex        =   59
         Top             =   240
         Width           =   825
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   19
         Left            =   10560
         TabIndex        =   40
         Top             =   240
         Width           =   540
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Guia:"
         Height          =   195
         Index           =   12
         Left            =   480
         TabIndex        =   39
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   5055
      Left            =   480
      TabIndex        =   24
      Top             =   1200
      Width           =   9015
      Begin VB.ComboBox CboTipo 
         Height          =   315
         ItemData        =   "FrmRemisiones.frx":0000
         Left            =   960
         List            =   "FrmRemisiones.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton CmdImportarDatos 
         Caption         =   "&Importar"
         Height          =   255
         Left            =   8160
         TabIndex        =   81
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Campo 
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
         Index           =   30
         Left            =   960
         MaxLength       =   5
         TabIndex        =   10
         Tag             =   "1"
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox Campo 
         Height          =   285
         Index           =   29
         Left            =   5040
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Campo 
         Height          =   285
         Index           =   28
         Left            =   960
         MaxLength       =   250
         TabIndex        =   15
         Top             =   4680
         Width           =   7935
      End
      Begin VB.CommandButton CmdBuscarNegociacion 
         Caption         =   "..."
         Height          =   255
         Left            =   8160
         TabIndex        =   70
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Campo 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   25
         Left            =   2520
         MaxLength       =   60
         TabIndex        =   69
         ToolTipText     =   "Aqui se debe ingresar el tercero"
         Top             =   480
         Width           =   5535
      End
      Begin VB.TextBox Campo 
         Height          =   285
         Index           =   27
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   68
         Tag             =   "1"
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox Campo 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   24
         Left            =   960
         MaxLength       =   11
         TabIndex        =   1
         ToolTipText     =   "Aqui se debe ingresar el tercero"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Campo 
         Height          =   285
         Index           =   9
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   44
         Tag             =   "1"
         Top             =   3240
         Width           =   975
      End
      Begin VB.CheckBox ChKCPorte 
         Alignment       =   1  'Right Justify
         Caption         =   "Devolver cartaporte"
         Height          =   255
         Left            =   7080
         TabIndex        =   6
         Top             =   1480
         Width           =   1815
      End
      Begin VB.TextBox Campo 
         Height          =   765
         Index           =   22
         Left            =   3120
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   3840
         Width           =   5775
      End
      Begin VB.TextBox Campo 
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
         Index           =   8
         Left            =   960
         MaxLength       =   5
         TabIndex        =   11
         Tag             =   "1"
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox Campo 
         Height          =   285
         Index           =   6
         Left            =   960
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2160
         Width           =   7935
      End
      Begin VB.TextBox Campo 
         Height          =   285
         Index           =   7
         Left            =   7680
         MaxLength       =   11
         TabIndex        =   8
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Campo 
         Height          =   285
         Index           =   5
         Left            =   960
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1800
         Width           =   5895
      End
      Begin VB.TextBox Campo 
         Height          =   285
         Index           =   4
         Left            =   6960
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Campo 
         Height          =   285
         Index           =   2
         Left            =   960
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1200
         Width           =   3615
      End
      Begin VB.ComboBox CboTpServicio 
         Height          =   315
         ItemData        =   "FrmRemisiones.frx":0047
         Left            =   960
         List            =   "FrmRemisiones.frx":0060
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3840
         Width           =   2055
      End
      Begin VB.TextBox Campo 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   960
         MaxLength       =   11
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cobro:"
         Height          =   195
         Left            =   360
         TabIndex        =   110
         Top             =   4200
         Width           =   465
      End
      Begin VB.Label LblDepartamentoDestino 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6360
         TabIndex        =   96
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label LblDepartamentoOrigen 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6360
         TabIndex        =   95
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Origen:"
         Height          =   195
         Index           =   13
         Left            =   405
         TabIndex        =   80
         Top             =   2520
         Width           =   510
      End
      Begin VB.Label LblConsulta 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   79
         Top             =   2520
         Width           =   4215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rel:"
         Height          =   195
         Left            =   4680
         TabIndex        =   78
         Top             =   1200
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ref:"
         Height          =   195
         Left            =   600
         TabIndex        =   71
         Top             =   4680
         Width           =   300
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta:"
         Height          =   195
         Index           =   33
         Left            =   360
         TabIndex        =   67
         Top             =   480
         Width           =   555
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Index           =   32
         Left            =   525
         TabIndex        =   46
         Top             =   3240
         Width           =   390
      End
      Begin VB.Label LblConsulta 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   45
         Top             =   3240
         Width           =   6255
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Id Neg:"
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   37
         Top             =   840
         Width           =   525
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   1
         Left            =   330
         TabIndex        =   36
         Top             =   2880
         Width           =   585
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   35
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   34
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Remitente:"
         Height          =   195
         Index           =   10
         Left            =   150
         TabIndex        =   33
         Top             =   1200
         Width           =   765
      End
      Begin VB.Line LnSeparadora 
         BorderWidth     =   3
         Index           =   2
         X1              =   960
         X2              =   8880
         Y1              =   3645
         Y2              =   3645
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   345
         TabIndex        =   32
         Top             =   3480
         Width           =   570
      End
      Begin VB.Label LblConsulta 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   31
         Top             =   2880
         Width           =   4215
      End
      Begin VB.Line LnSeparadora 
         BorderWidth     =   3
         Index           =   1
         X1              =   1200
         X2              =   6960
         Y1              =   1605
         Y2              =   1605
      End
      Begin VB.Line LnSeparadora 
         BorderWidth     =   3
         Index           =   0
         X1              =   960
         X2              =   8880
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label LblConsulta 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   30
         Top             =   840
         Width           =   5535
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Servicio:"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   29
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Doc:"
         Height          =   195
         Index           =   7
         Left            =   6600
         TabIndex        =   28
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   180
         TabIndex        =   27
         Top             =   180
         Width           =   735
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   26
         Top             =   1500
         Width           =   810
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Telefono:"
         Height          =   195
         Index           =   3
         Left            =   6960
         TabIndex        =   25
         Top             =   1800
         Width           =   675
      End
   End
   Begin MSComctlLib.ListView LstTem 
      Height          =   1935
      Left            =   480
      TabIndex        =   47
      Top             =   6360
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   3413
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
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Producto"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ID"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Empaque"
         Object.Width           =   4057
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
         Object.Width           =   1411
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
         Object.Width           =   1940
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolRemisiones 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Width           =   16680
      _ExtentX        =   29422
      _ExtentY        =   1005
      ButtonWidth     =   847
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuev"
            Object.ToolTipText     =   "Crear nuevo registro [F9]"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Guar"
            Object.ToolTipText     =   "Guarda la informacio [F11]"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Object.ToolTipText     =   "Editar la informacion guardada [F10]"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Elim"
            Object.ToolTipText     =   "Elimina o anula el registro [F3]"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Can"
            Object.ToolTipText     =   "Cancela la creacion del registro [F4]"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Bus"
            Object.ToolTipText     =   "Buscar [Inicio]"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "Buscar1"
                  Text            =   "Por documento (Ctr+D)"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pri"
            Object.ToolTipText     =   "Ir al primer registro [F5]"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ant"
            Object.ToolTipText     =   "Ir al anterior registro [F6]"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sig"
            Object.ToolTipText     =   "Ir al siguiente registro [F7]"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ult"
            Object.ToolTipText     =   "Ir al ultimo registro [F8]"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cer"
            Object.ToolTipText     =   "Cerrar esta ventana [F12]"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Act"
            Object.ToolTipText     =   "Actualizar la informacion"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imp"
            Object.ToolTipText     =   "Imprimir registro [Fin]"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Car"
            Object.ToolTipText     =   "Carga informacion adicional [Pausa]"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Acc"
            Style           =   5
            Object.Width           =   1e-4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "Accion1"
                  Text            =   "Quitar estado de impreso"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "Accion3"
                  Text            =   "Reimprimir guia"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "Accion2"
                  Text            =   "Cambiar numero de guia"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmRemisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Editando As Boolean
Dim ClienteViejo As String
Dim COViejo As Integer
Dim NegociacionVieja As Double
Dim GuiaConsecutivo As Long
Dim UltRemision As Long
Dim GuiaFormato As Boolean

Dim TpServicios(7) As Byte
Dim CPorte As Byte
Dim liquidado As Boolean
Dim PermiteRecaudo As Boolean
Dim NegociacionInactiva As Boolean
Dim ListaPreciosVencida As Boolean
Dim rstGuias As New ADODB.Recordset
Dim strSqlGuias As String
Dim douPorcentajeManejo As Double
Dim douMinimoManejoUnidad As Double
Dim douMinimoManejoDespacho As Double
Dim intIdListaPrecios As Integer
Dim douVrKilo As Double
Dim douDctoKilo As Double
Dim douKilosMinimos As Double
Dim boolNoAplicarDctoReexpediciones As Integer
Dim boolRedondearFlete As Integer
Dim ManejaCobroContado As Integer
Dim ManejaCobroDestino As Integer
Dim ManejaCobroCorriente As Integer

Private Sub Campo_Change(Index As Integer)
  If Index = 19 Then TxtEstado.Text = DevEstadoDespacho(Campo(19))
End Sub

Private Sub Campo_GotFocus(Index As Integer)
  EnfocarT Campo(Index)
  Campo(Index).BackColor = &H80000001
  Campo(Index).ForeColor = &HFFFFFF
End Sub
Private Sub Campo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Select Case Index
      Case 24
        Principal.ToolConsultas1.AbrirDevConsulta 7, CnnPrincipal
        Campo(24).Text = Principal.ToolConsultas1.DatSt
  
      Case 8
        Principal.ToolConsultas1.AbrirDevConsulta 1, CnnPrincipal
        Campo(8).Text = Principal.ToolConsultas1.DatLo
    
      Case 30
        Principal.ToolConsultas1.AbrirDevConsulta 1, CnnPrincipal
        Campo(30).Text = Principal.ToolConsultas1.DatLo
    
      Case 2
        Principal.ToolConsultas1.AbrirDevConsulta 4, CnnPrincipal
        If Principal.ToolConsultas1.DatSt <> "" Then
          AbrirRecorset rstUniversal, "Select*From Remitentes Where IdRemitente='" & Principal.ToolConsultas1.DatSt & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
            If rstUniversal.EOF = False Then
              Campo(2) = rstUniversal!NmRemitente & ""
            Else
              Campo(2) = ""
            End If
          CerrarRecorset rstUniversal
        End If
        
      Case 5
        Principal.ToolConsultas1.AbrirDevConsulta 9, CnnPrincipal
        FufuSt = Principal.ToolConsultas1.DatSt
        If FufuSt <> "" Then
          AbrirRecorset rstUniversal, "Select*From Destinatarios Where IdDestinatario='" & FufuSt & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
            If rstUniversal.EOF = False Then
              Campo(5) = rstUniversal!NmDestinatario & ""
              Campo(6) = rstUniversal!DirDestinatario & ""
              Campo(7) = rstUniversal!TelDestinatario & ""
              Campo(8) = rstUniversal!IdCiuDestinatario & ""
            Else
              Campo(5) = ""
              Campo(6) = ""
              Campo(7) = ""
              Campo(8) = ""
            End If
          CerrarRecorset rstUniversal
        End If
    Case 12, 13, 14, 15, 16, 17, 18
      Liquidacion
    
    End Select
  End If
End Sub

Private Sub Campo_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
    Case 7, 8, 26, 30, 12, 13, 14, 15, 16, 17, 18
      ValidarEntrada Campo(Index), KeyAscii, 1
      If Index = 26 Then
        If PermiteRecaudo = False And KeyAscii <> 13 Then
          MsgBox "Este cliente no permite recaudo, para cambiar esta opcion", vbCritical
          KeyAscii = 0
          Campo(26) = 0
        End If
      End If
    Case 22
      If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys vbTab
      End If
  End Select
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If
End Sub

Private Sub Campo_LostFocus(Index As Integer)
  Campo(Index).BackColor = &H80000005
  Campo(Index).ForeColor = &H80000012
End Sub

Private Sub Campo_Validate(Index As Integer, Cancel As Boolean)
  Select Case Index
    Case 8
      If Val(Campo(8).Text) <> 0 Then
        AbrirRecorset rstUniversal, "SELECT IdCiudad, NmCiudad, NmDepartamento FROM Ciudades LEFT JOIN departamentos ON ciudades.IdDepartamento = departamentos.IdDepartamento WHERE IdCiudad = " & Campo(8), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          LblConsulta(1) = rstUniversal!NmCiudad & ""
          LblDepartamentoDestino.Caption = rstUniversal!NmDepartamento & ""
          CerrarRecorset rstUniversal
          AbrirRecorset rstUniversal, "Select* from Rutas_Ciudades where IdCiudad=" & Campo(8), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
          If rstUniversal.EOF = False Then
            Campo(9).Text = rstUniversal.Fields("IdRuta")
            Campo(27).Text = rstUniversal.Fields("Orden")
          Else
            Campo(9) = 1
            Campo(27) = 0
            LblConsulta(2).Caption = ""
          End If
          CerrarRecorset rstUniversal
          If Val(Campo(9).Text) <> 0 Then
            AbrirRecorset rstUniversal, "SELECT IdRuta, NmRuta From Rutas where IdRuta=" & Campo(9), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
            If rstUniversal.EOF = False Then
              LblConsulta(2) = rstUniversal!NmRuta & ""
            Else
              Campo(9) = ""
              LblConsulta(2).Caption = ""
            End If
          End If
        Else
          LblConsulta(1) = "": Campo(8) = "": Campo(9) = ""
        End If
        CerrarRecorset rstUniversal
      End If
    Case 24
      If Campo(24).Text <> "" Then
        AbrirRecorset rstUniversal, "Select IdTercero, RazonSocial, IdCliente, ManejaCobroContado, ManejaCobroDestino, ManejaCobroCorriente, Inactivo from Terceros where IdTercero='" & Campo(24) & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
          If rstUniversal.EOF = False Then
            If Val(rstUniversal.Fields("Inactivo")) = 0 Then
              Campo(25).Text = rstUniversal.Fields("RazonSocial") & ""
              ManejaCobroContado = rstUniversal.Fields("ManejaCobroContado")
              ManejaCobroDestino = rstUniversal.Fields("ManejaCobroDestino")
              ManejaCobroCorriente = rstUniversal.Fields("ManejaCobroCorriente")
              If Campo(2) = "" Then Campo(2) = rstUniversal.Fields("RazonSocial") & ""
              Campo(3).Text = rstUniversal.Fields("IdCliente") & ""
              If ClienteViejo = Campo(24) Then
                Campo(3).Text = NegociacionVieja
              End If
              CerrarRecorset rstUniversal
              CargarNegociacion
            Else
              MsgBox "El cliente se encuentra inactivo", vbCritical
              Campo(24).Text = ""
            End If
          Else
            If MsgBox("El tercero no existe ¿Desea crearlo?", vbQuestion + vbYesNo, "No existe el tercero") = vbYes Then
              FufuSt = Campo(24).Text
              FrmAgregarTercero.Show 1
              Cancel = True
              Campo(24).Text = FufuSt
            Else
              Campo(24).Text = ""
            End If
          End If
        CerrarRecorset rstUniversal
      End If
      
    Case 30
      If Val(Campo(30).Text) <> 0 Then
        AbrirRecorset rstUniversal, "SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Campo(30), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          LblConsulta(3) = rstUniversal!NmCiudad & ""
        Else
          LblConsulta(3) = "": Campo(30) = ""
        End If
        CerrarRecorset rstUniversal
      End If
      
    Case 12
      If Val(Campo(14).Text) = 0 And Val(Campo(12).Text) > 0 Then
        Campo(14).Text = Val(Campo(12).Text) * douPorcentajeManejo / 100
        If douMinimoManejoDespacho > Val(Campo(14).Text) Then
          Campo(14).Text = douMinimoManejoDespacho
        End If
        If (douMinimoManejoUnidad * Val(Campo(15).Text)) > Val(Campo(14).Text) Then
          Campo(14).Text = douMinimoManejoUnidad * Val(Campo(15).Text)
        End If
      End If
      
    Case 15
      AbrirRecorset rstUniversal, "SELECT Minimos FROM listaspreciosciudades WHERE IdListaPrecios = " & intIdListaPrecios & " AND IdCiudadOrigen = " & Val(Campo(30).Text) & " AND IdCiudad = " & Val(Campo(8).Text) & " AND IdProducto = 1 AND VrKilo > 0", CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.RecordCount > 0 Then
        If Val(Campo(17).Text) <= 0 Then
          If Val(rstUniversal!Minimos) > 0 Then
            Campo(17).Text = Val(rstUniversal!Minimos) * Val(Campo(15).Text)
          End If
        End If
      End If
      If douKilosMinimos > 0 Then
        Campo(17).Text = douKilosMinimos * Val(Campo(15).Text)
        If Val(Campo(16).Text) <= 0 Then
          Campo(16).Text = douKilosMinimos * Val(Campo(15).Text)
        End If
      End If
      
      CerrarRecorset rstUniversal
      
    Case 16
      LiquidarKilosFacturar
      
    Case 17
      If Val(Campo(13).Text) = 0 And Val(Campo(8).Text) <> 0 Then
        Dim rstListaPreciosDetalle As New ADODB.Recordset
        rstListaPreciosDetalle.CursorLocation = adUseClient
        AbrirRecorset rstListaPreciosDetalle, "SELECT VrKilo, Minimos FROM listaspreciosciudades WHERE IdListaPrecios = " & intIdListaPrecios & " AND IdCiudadOrigen = " & Val(Campo(30).Text) & " AND IdCiudad = " & Val(Campo(8).Text) & " AND IdProducto = 1 AND VrKilo > 0", CnnPrincipal, adOpenDynamic, adLockOptimistic
        If rstListaPreciosDetalle.RecordCount > 0 Then
          douVrKilo = rstListaPreciosDetalle!VrKilo
          AbrirRecorset rstUniversal, "SELECT Reexpedicion FROM ciudades WHERE Reexpedicion = 1 AND IdCiudad = " & Val(Campo(8).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
          If rstUniversal.RecordCount > 0 Then
            If boolNoAplicarDctoReexpediciones = 1 Then
              douDctoKilo = 0
            End If
          End If
          Dim douFlete As Double
          douFlete = Val(Campo(17).Text) * (douVrKilo - (douVrKilo * douDctoKilo / 100))
          If boolRedondearFlete = 1 Then
            douFlete = Round(douFlete * 0.01) * 100
          Else
            douFlete = Round(douFlete)
          End If
          Campo(13).Text = douFlete
        End If
      End If
      
    Case 18
      LiquidarKilosFacturar
  End Select
End Sub

Private Sub CboTipo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub CboTpServicio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub CboTpServicio_Validate(Cancel As Boolean)
  If TpServicios(CboTpServicio.ListIndex + 1) = False Then MsgTit "Este tipo de servico no esta permitido por el cliente": CboTpServicio.ListIndex = -1
End Sub
Private Sub ChKCPorte_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub ChkGuiFac_KeyPress(KeyAscii As Integer)
  'If KeyAscii = 13 Then Liquidacion
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub CmdBuscarNegociacion_Click()
  If Campo(24).Text <> "" Then
    FufuSt = Campo(24).Text
    FrmBuscarNegociaciones.Show 1
    Campo(2).SetFocus
    If FufuLo <> 0 Then
      Campo(3).Text = FufuLo
      CargarNegociacion
    End If
    
  End If
End Sub

Private Sub CmdBuscarPorDocumento_Click()
  BuscarPorDocumento
End Sub

Private Sub CmdCambiarCO_Click()
  FrmBuscarCO.Show 1
  If FufuLo <> 0 Then
    Campo(23).Text = FufuLo
    If Val(Campo(23).Text) <> 0 Then TxtCOIng.Text = DevResBus("SELECT IdPO, NmPuntoOperaciones From CentrosOperaciones where IdPO=" & Campo(23), "NmPuntoOperaciones", CnnPrincipal)
  End If
End Sub



Private Sub CmdGenerarRecibo_Click()
  AbrirRecorset rstUniversal, "SELECT Guia, GuiFac, Estado, Anulada FROM guias WHERE Guia = " & Val(Campo(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.RecordCount > 0 Then
    If Val(rstUniversal!GuiFac) = 1 Then
      If Val(rstUniversal!Anulada) = 1 Then
        MsgBox "No se le pueden generar recibos a una guia que se encuentra anulada"
      Else
        If rstUniversal!Estado = "I" Or rstUniversal.Fields("Estado") = "G" Or rstUniversal.Fields("Estado") = "P" Then
          FufuLo = Val(Campo(0).Text)
          FrmGenerarReciboCaja.Show 1
        Else
          If MsgBox("La guia debe estar impresa, desea imprimir la guia? si responde si, se le asignara el estado de impresa pero no se enviara el formato a la impresora", vbQuestion + vbYesNo) = vbYes Then
            AbrirRecorset rstUniversal, "Update Guias Set Estado='I' where Guia=" & Val(Campo(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
            Campo(19) = "I"
            AccionTool 17
            CmdGenerarRecibo_Click
          End If
        End If
      End If
    Else
      MsgBox "Solo se le pueden realizar recibos a las guias contado y contraentrega (Facturas)", vbCritical
    End If
  End If
  CerrarRecorset rstUniversal
End Sub

Private Sub CmdImportarDatos_Click()
    If Principal.ToolConsultas1.AbrirDevDatos("Numero de documento", "Digite el numero del documento", 2, 0) = True Then
      'MsgBox DevDocSinCeros(Principal.ToolConsultas1.DatSt)
      AbrirRecorset rstUniversal, "Select*from guias_imp where Documento='" & DevDocSinCeros(Principal.ToolConsultas1.DatSt) & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        Dim rstCiudad As New ADODB.Recordset
        rstCiudad.CursorLocation = adUseClient
        If (rstUniversal.Fields("IdDestino") & "") <> "" Then
          AbrirRecorset rstCiudad, "Select IdCiudad from ciudades where CodMinTrans='" & rstUniversal.Fields("IdDestino") & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
            If rstCiudad.RecordCount > 0 Then Campo(8).Text = rstCiudad.Fields("IdCiudad")
          CerrarRecorset rstCiudad
        End If

        Campo(2).Text = rstUniversal.Fields("Remitente") & ""
        Campo(29).Text = rstUniversal.Fields("Relacion") & ""
        Campo(4).Text = rstUniversal.Fields("Documento") & ""
        Campo(5).Text = rstUniversal.Fields("NmDestinatario") & ""
        Campo(6).Text = rstUniversal.Fields("DirDestinatario") & ""
        Campo(7).Text = rstUniversal.Fields("TelDestinatario") & ""
        Campo(22).Text = rstUniversal.Fields("Observaciones") & ""
        Campo(12).Text = Val(rstUniversal.Fields("Declarado") & "")
        
        CboTpServicio.ListIndex = 0
        Campo(30).SetFocus
      Else
        MsgBox "No se encontraron importaciones con este documento", vbCritical
      End If
      CerrarRecorset rstUniversal
    End If
End Sub

Private Sub CmdVerProductos_Click()
  LstTem.ListItems.Clear
  AbrirRecorset rstUniversal, "SELECT MvtoGuias.*, Productos.NmProducto, Empaques.NmEmpaque FROM MvtoGuias INNER JOIN Productos ON MvtoGuias.IdProducto = Productos.IdProducto INNER JOIN Empaques ON MvtoGuias.IdEmpaque = Empaques.IdEmpaque Where MvtoGuias.Guia = " & Campo(0), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    Set Item = LstTem.ListItems.Add(, , rstUniversal.Fields("Lote") & "")
        Item.SubItems(1) = rstUniversal.Fields("IdProducto")
        Item.SubItems(2) = rstUniversal.Fields("NmProducto")
        Item.SubItems(3) = rstUniversal.Fields("IdEmpaque")
        Item.SubItems(4) = rstUniversal.Fields("NmEmpaque")
        Item.SubItems(5) = rstUniversal.Fields("Ancho")
        Item.SubItems(6) = rstUniversal.Fields("Largo")
        Item.SubItems(7) = rstUniversal.Fields("Altura")
        Item.SubItems(8) = rstUniversal.Fields("Cant")
        Item.SubItems(9) = rstUniversal.Fields("KilosReal")
        Item.SubItems(10) = rstUniversal.Fields("KilosVol")
        Item.SubItems(11) = rstUniversal.Fields("KilosFacturados")
        Item.SubItems(12) = rstUniversal.Fields("VlrFlete")
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  LaTecla (KeyCode), ToolRemisiones
  If KeyCode = 65 And Shift = 2 Then
    BuscarPorDocumento
  End If
End Sub

Private Sub Form_Load()
  IconosTool ToolRemisiones, Principal.IgListTool
  rstGuias.CursorLocation = adUseClient
  strSqlGuias = "SELECT guias.* " & _
                "FROM guias "

  AbrirRecorset rstGuias, strSqlGuias & " ORDER BY FhEntradaBodega DESC LIMIT 20", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Formatos rstGuias
  Asignar rstGuias
  COViejo = Coperaciones
  GuiaConsecutivo = 0
  GuiaFormato = DevImprimirGuiaFormato
  
  iniciarVariablesLiquidacion
End Sub
Private Sub Formatos(rstForma As ADODB.Recordset)
  For II = 0 To 35
    Set rstForma.Fields(II).DataFormat = Campo(II).DataFormat
  Next
End Sub

Private Sub Asignar(rstAsignar As ADODB.Recordset)
  For II = 0 To 35
    Campo(II).Text = rstAsignar.Fields(II) & ""
  Next
  CboTpServicio.ListIndex = Val(rstAsignar!TpServicio)
  CboTipo.ListIndex = Val(rstAsignar!GuiaTipo) - 1
  ChKCPorte.value = DevCheck(rstAsignar!CPorte)
  
  ChkDespachada.value = DevCheck(rstAsignar!Despachada)
  ChkEntregada.value = DevCheck(rstAsignar!Entregada)
  ChkDescargada.value = DevCheck(rstAsignar!Descargada)
  ChkAnulada.value = DevCheck(rstAsignar!Anulada)
  ChkFacturada.value = DevCheck(rstAsignar!Facturada)
  ChkGuiFac.value = DevCheck(rstAsignar!GuiFac)
  ChkRelacionada.value = DevCheck(rstAsignar!Relacionada)
  ChkNovedad.value = DevCheck(rstAsignar!EnNovedad)
  
  
  LblConsulta(0).Caption = ""
  LstTem.ListItems.Clear
  TxtCOIng.Text = ""
  TxtCOCar.Text = ""
  
  LblNmUsuario.Caption = DevResBus("SELECT IDUsuario, NmUsuario From usuarios where IDUsuario=" & rstAsignar.Fields("IdUsuario"), "NmUsuario", CnnPrincipal)
  Dim strSql As String
  Dim rstCiudades As New ADODB.Recordset
  rstCiudades.CursorLocation = adUseClient
  strSql = "SELECT NmCiudad, NmDepartamento FROM ciudades LEFT JOIN departamentos ON ciudades.IdDepartamento = departamentos.IdDepartamento WHERE IdCiudad=" & Campo(30)
  AbrirRecorset rstCiudades, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstCiudades.RecordCount > 0 Then
    LblConsulta(3).Caption = rstCiudades.Fields("NmCiudad") & ""
    LblDepartamentoOrigen.Caption = rstCiudades.Fields("NmDepartamento") & ""
  End If
  CerrarRecorset rstCiudades
  
  strSql = "SELECT NmCiudad, NmDepartamento FROM ciudades LEFT JOIN departamentos ON ciudades.IdDepartamento = departamentos.IdDepartamento WHERE IdCiudad=" & Campo(8)
  AbrirRecorset rstCiudades, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstCiudades.RecordCount > 0 Then
    LblConsulta(1).Caption = rstCiudades.Fields("NmCiudad") & ""
    LblDepartamentoDestino.Caption = rstCiudades.Fields("NmDepartamento") & ""
  End If
  CerrarRecorset rstCiudades
    
  LblConsulta(2).Caption = DevResBus("SELECT IdRuta, NmRuta From rutas where IdRuta=" & Campo(9), "NmRuta", CnnPrincipal)
  TxtCOIng.Text = DevResBus("SELECT IdPO, NmPuntoOperaciones From CentrosOperaciones where IdPO=" & Campo(23), "NmPuntoOperaciones", CnnPrincipal)
  TxtCOCar.Text = DevResBus("SELECT IdPO, NmPuntoOperaciones From CentrosOperaciones where IdPO=" & Campo(1), "NmPuntoOperaciones", CnnPrincipal)
End Sub
Private Sub limpiar()
For II = 0 To 35
  Campo(II).Text = ""
Next
LblConsulta(0).Caption = ""
LblConsulta(1).Caption = ""
LblConsulta(2).Caption = ""
LblConsulta(3).Caption = ""
LblDepartamentoOrigen.Caption = ""
LblDepartamentoDestino.Caption = ""
CboTpServicio.ListIndex = -1
ChKCPorte.value = 0
ChkDespachada.value = 0
ChkEntregada.value = 0
ChkDescargada.value = 0
ChkAnulada.value = 0
ChkFacturada.value = 0
ChkRelacionada.value = 0
ChkNovedad.value = 0
LstTem.ListItems.Clear
TxtCOIng.Text = ""
TxtCOCar.Text = ""
LblNmUsuario.Caption = UsuarioActivo
End Sub
Private Sub Desbloquear()
  FraDatos.Enabled = True
  FraLiquidacion.Enabled = True
  'LblMensajeLiquidacion.Visible = True
  CmdCambiarCO.Enabled = True
  CmdVerProductos.Enabled = False
  BotTool 3, 17, ToolRemisiones, True
End Sub
Private Sub Bloquear()
  FraDatos.Enabled = False
  CmdCambiarCO.Enabled = False
  FraLiquidacion.Enabled = False
  'LblMensajeLiquidacion.Visible = False
  CmdVerProductos.Enabled = True
  BotTool 3, 17, ToolRemisiones, False
End Sub
Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
      If CpPermiso(1, CodUsuarioActivo, 2, CnnPrincipal) = True Then
        Desbloquear
        liquidado = False
        limpiar
        iniciarVariablesLiquidacion
        Campo(1) = Coperaciones
        Campo(23) = COViejo
        If Val(Campo(23).Text) <> 0 Then TxtCOIng.Text = DevResBus("SELECT IdPO, NmPuntoOperaciones From CentrosOperaciones where IdPO=" & Campo(23), "NmPuntoOperaciones", CnnPrincipal)
        If Val(Campo(23).Text) <> 0 Then Campo(30).Text = DevResBus("SELECT IdPO, IdCiudad From CentrosOperaciones where IdPO=" & Campo(23), "IdCiudad", CnnPrincipal)
        If Val(Campo(30).Text) <> 0 Then LblConsulta(3).Caption = DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Campo(30), "NmCiudad", CnnPrincipal)
        If Val(Campo(1).Text) <> 0 Then TxtCOCar.Text = DevResBus("SELECT IdPO, NmPuntoOperaciones From CentrosOperaciones where IdPO=" & Campo(1), "NmPuntoOperaciones", CnnPrincipal)
        Campo(10) = Date
        Campo(11) = Date - 1
        Campo(24).Text = ClienteViejo
        CboTpServicio.ListIndex = 0
        CboTipo.ListIndex = 0
        Campo(24).SetFocus
      End If
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            AbrirRecorset rstUniversal, "Update Guias set GuiFac=" & ChkGuiFac.value & ", CR=" & Campo(1) & ", COIng=" & Val(Campo(23)) & ", Remitente='" & Campo(2) & "', Cliente='" & Campo(25) & "', Cuenta='" & Campo(24) & "', " & _
            "IdCliente='" & Campo(3) & "', DocCliente='" & Campo(4) & "', EmpaqueRef='" & Campo(28) & "', RelCliente='" & Campo(29) & "', IdCiuOrigen=" & Campo(30) & ", NmDestinatario='" & Campo(5) & "', DirDestinatario='" & Campo(6) & "', TelDestinatario='" & Campo(7) & "', IdCiuDestino=" & Campo(8) & ", IdRuta=" & Campo(9) & ", VrDeclarado=" & Campo(12) & ", VrFlete=" & Val(Campo(13)) & ", VrManejo=" & Val(Campo(14)) & ", Unidades=" & Campo(15) & ", KilosReales=" & Campo(16) & ", KilosFacturados=" & Campo(17) & ", KilosVolumen=" & Campo(18) & ", Observaciones='" & Campo(22) & "', Recaudo=" & Val(Campo(26)) & ", TpServicio=" & CboTpServicio.ListIndex & " where Guia=" & Val(Campo(0)), CnnPrincipal, adOpenDynamic, adLockOptimistic
            InsertarLog 6, Val(Campo(0).Text)
            VaciarGrilla
            Editando = False
            Bloquear
            AccionTool 17
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          If DevTipoCobro(Val(CboTipo.ListIndex) + 1) = 1 Or DevTipoCobro(Val(CboTipo.ListIndex) + 1) = 2 Then
            If DevConsecutivoGuiasFactura = True Then
              FufuLo = SacarConsecutivo("GuiasFactura", CnnPrincipal)
            Else
              If Principal.ToolConsultas1.AbrirDevDatos("Digite el numero de la guia", "Digite el numero para la guia que esta digitando", 3, GuiaConsecutivo) = True Then
                FufuLo = Principal.ToolConsultas1.DatLo
              Else
                MsgBox "Debe digitar un numero de guia", vbCritical
                FufuLo = 0
              End If
            End If
          Else
            If GuiaManConsecutivo = True Then
              FufuLo = SacarConsecutivo("Guias", CnnPrincipal)
            Else
              If Principal.ToolConsultas1.AbrirDevDatos("Digite el numero de la guia", "Digite el numero para la guia que esta digitando", 3, GuiaConsecutivo) = True Then
                FufuLo = Principal.ToolConsultas1.DatLo
              Else
                MsgBox "Debe digitar un numero de guia", vbCritical
                FufuLo = 0
              End If
            End If
          End If
            
          If FufuLo <> 0 Then
            If ComprobarExGuia(FufuLo, Val(LblGuiaTipo.Tag)) = False Then
              Campo(0).Text = FufuLo
              Campo(10).Text = Format(DevFechaEspecial(), "yyyy-mm-dd HH:mm:ss")
              AbrirRecorset rstUniversal, "INSERT INTO Guias " & _
              "(Guia, CR, Remitente, IdCliente, DocCliente, NmDestinatario, DirDestinatario, TelDestinatario, IdCiuDestino, IdRuta, " & _
              "FhEntradaBodega, VrDeclarado, VrFlete, VrManejo, Unidades, KilosReales, KilosFacturados, KilosVolumen, " & _
              "Estado, IdFactura, IdDespacho, Observaciones, COIng, Cuenta, Cliente, Recaudo, orden, EmpaqueRef, RelCliente, IdCiuOrigen, TpServicio, CPorte, Entregada, Descargada, Despachada, Anulada, GuiFac, Facturada, IdUsuario, IdEmpresa, GuiaTipo, TipoCobro) " & _
              "VALUES(" & Campo(0) & "," & Campo(1) & ",'" & Campo(2) & "', '" & Campo(3) & "','" & Campo(4) & "','" & Campo(5) & "','" & Campo(6) & "','" & Campo(7) & "', " & Val(Campo(8)) & ", " & Val(Campo(9)) & ", " & _
              "'" & Campo(10).Text & "', " & Val(Campo(12)) & ", " & Val(Campo(13)) & ", " & Val(Campo(14)) & ", " & Val(Campo(15)) & ", " & Val(Campo(16)) & ", " & Val(Campo(17)) & ", " & Val(Campo(18)) & ", " & _
              "'D', 0, null, '" & Campo(22) & "', " & Val(Campo(23)) & ", '" & Campo(24) & "', '" & Campo(25) & "', " & Val(Campo(26)) & ", " & Val(Campo(27)) & ", '" & Campo(28).Text & "','" & Campo(29).Text & "'," & Campo(30).Text & ", " & CboTpServicio.ListIndex & ", " & ChKCPorte.value & ", 0, 0, 0, 0, " & DevTpGuiaFactura(Val(CboTipo.ListIndex) + 1) & ", 0, " & CodUsuarioActivo & ", " & CodEmpresaActiva & ", " & Val(CboTipo.ListIndex) + 1 & ", " & DevTipoCobro(Val(CboTipo.ListIndex) + 1) & ")", CnnPrincipal, adOpenDynamic, adLockBatchOptimistic
              InsertarLog 1, Val(Campo(0).Text)
              GuiaConsecutivo = FufuLo + 1
              VaciarGrilla
              Bloquear
              Campo(19) = "D"
              ClienteViejo = Campo(24).Text
              COViejo = Val(Campo(23).Text)
              NegociacionVieja = Val(Campo(3).Text)
              UltRemision = FufuLo
              CmdVerProductos.SetFocus
            Else
              MsgBox "Esta guia ya existe, pruebe con otro numero de guia", vbCritical, "El numero de guia ya existe"
              AccionTool 4
            End If
          End If
        End If
      End If
    Case 5  'Editar
      If CpPermiso(1, CodUsuarioActivo, 3, CnnPrincipal) = True Then
      AbrirRecorset rstUniversal, "Select Guia, Facturada, Anulada FROM guias WHERE Guia=" & Val(Campo(0).Text) & " and Facturada=0 and Anulada=0 and Estado='D'", CnnPrincipal, adOpenDynamic, adLockOptimistic
        FufuLo = rstUniversal.RecordCount
      CerrarRecorset rstUniversal
      
      If FufuLo >= 1 Then
        Editando = True
        liquidado = True
        AbrirRecorset rstUniversal, strSqlGuias & " WHERE Guia=" & Campo(0), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        Asignar rstUniversal
        Desbloquear
        LlenarGrillaProductos
        iniciarVariablesLiquidacion
        IdClienteViejo = Campo(3).Text
        Campo(24).SetFocus
      Else
        MsgBox "Esta guia no se puede editar porque esta facturada o anulada o impresa, si desea puede quitarle el estado de impreso a la guia", vbCritical, "No se puede editar"
      End If
      End If
    Case 6 'Eliminar
      If CpPermiso(1, CodUsuarioActivo, 4, CnnPrincipal) = True Then
        AbrirRecorset rstUniversal, "Select Guia, IdFactura, Anulada, GuiFac, Estado, Despachada FROM guias WHERE Guia = " & Val(Campo(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
        If rstUniversal.RecordCount > 0 Then
          If rstUniversal.Fields("IdFactura") = 0 And Val(rstUniversal.Fields("Anulada")) = 0 Then
            If rstUniversal!Estado = "I" And Val(rstUniversal!Despachada) = 0 Then
              If MsgBox("Las guias no se pueden eliminar, solamente anular, ¿Esta seguro de anular la guia?", vbQuestion + vbYesNo, "Anular guias") = vbYes Then
                AbrirRecorset rstUniversal, "Update Guias set VrDeclarado=0, VrFlete=0, VrManejo=0, Abonos=0, Unidades=0, KilosReales=0, KilosFacturados=0, KilosVolumen=0, Estado='A', Anulada=1 where Guia=" & Campo(0).Text, CnnPrincipal, adOpenDynamic, adLockReadOnly
                AbrirRecorset rstUniversal, "Update recibos_caja_soporte set VrFlete=0, VrManejo=0, ValorTotal=0 where Guia=" & Campo(0).Text, CnnPrincipal, adOpenDynamic, adLockReadOnly
                rstGuias.Requery
                InsertarLog 2, Val(Campo(0).Text)
                Campo(19).Text = "A"
                For II = 12 To 18
                  Campo(II).Text = 0
                Next
              End If
            Else
              MsgBox "Solo se pueden anular las guias con estado [DIGITADO] y que no esten ni de viaje ni en reparto", vbCritical, "Imporsible anular la guia"
            End If
          Else
            MsgBox "No se pueden anular guias ya anuladas o facturadas", vbCritical
          End If
        End If
        CerrarRecorset rstUniversal
      End If
      
    Case 7 'Cancelar
      If Editando = True Then
        AccionTool 4
      Else
        Asignar rstGuias
        Bloquear
        Editando = False
      End If
    Case 9  'Buscar
      If Principal.ToolConsultas1.AbrirDevDatos("Numero de Guia", "Digite el numero de la guia que desea buscar", 3, 0) = True Then
        AbrirRecorset rstUniversal, strSqlGuias & " WHERE Guia=" & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          Asignar rstUniversal
        Else
          MsgBox "No se encontraron guias con este numero", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If
    Case 11 'Primero
      UPrimero rstGuias
      Asignar rstGuias
    Case 12 'Anterior
      UAnterior rstGuias
      Asignar rstGuias
    Case 13 'Siguiente
      USiguiente rstGuias
      Asignar rstGuias
    Case 14 'Ultimo
      UUltimo rstGuias
      Asignar rstGuias
    Case 16 'Cerrar
      CerrarRecorset rstGuias
      Unload Me
    Case 17 'Actualizar
      rstGuias.Requery
    Case 18 'Imprimir
      If ComprobarEstado(Val(Campo(0))) = "D" Then
        If GuiaFormato = True Then
          AbrirRecorset rstUniversal, "Update Guias Set Estado='I' where Guia=" & Val(Campo(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
          Mostrar_Reporte CnnPrincipal, 15, "Select*from sql_im_impguia where Guia=" & Val(Campo(0).Text), "", 2
          InsertarLog 7, Val(Campo(0).Text)
          Campo(19) = "I"
          AccionTool 17
        Else
          FufuLo = SelectForm("Rem Cuartas", Me.hwnd)
          If ImprimirGuia(Campo(0).Text) = True Then
            AbrirRecorset rstUniversal, "Update Guias Set Estado='I' where Guia=" & Val(Campo(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
            InsertarLog 7, Val(Campo(0).Text)
            Campo(19) = "I"
            AccionTool 17
            MsgBox "Se imprimio la guia con exito"
          End If
          establecerPapel
        End If
      Else
        MsgBox "La guia no esta en estado digitada", vbCritical
      End If
    Case 19 'Recargar
      If Campo(3).Text <> "" Then
        AbrirRecorset rstUniversal, "Select Id, NmNegociacion from Negociaciones where Id=" & Val(Campo(3)), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
          If rstUniversal.EOF = False Then
            LblConsulta(0).Caption = rstUniversal.Fields("NmNegociacion")
          End If
        CerrarRecorset rstUniversal
      End If
      
      If Val(Campo(8).Text) <> 0 Then LblConsulta(1).Caption = DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Campo(8), "NmCiudad", CnnPrincipal)
      If Val(Campo(9).Text) <> 0 Then LblConsulta(2).Caption = DevResBus("SELECT IdRuta, NmRuta From Rutas where IdRuta=" & Campo(9), "NmRuta", CnnPrincipal)
      If Val(Campo(30).Text) <> 0 Then LblConsulta(3).Caption = DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Campo(30), "NmCiudad", CnnPrincipal)
      If Val(Campo(23).Text) <> 0 Then TxtCOIng.Text = DevResBus("SELECT IdPO, NmPuntoOperaciones From CentrosOperaciones where IdPO=" & Campo(23), "NmPuntoOperaciones", CnnPrincipal)
      If Val(Campo(1).Text) <> 0 Then TxtCOCar.Text = DevResBus("SELECT IdPO, NmPuntoOperaciones From CentrosOperaciones where IdPO=" & Campo(1), "NmPuntoOperaciones", CnnPrincipal)
      FufuLo = Campo(0)
      FrmInfoGuia.Show 1
  End Select
End Sub



Private Sub Form_Unload(Cancel As Integer)
  Set rstGuias = Nothing
End Sub

Private Sub ToolRemisiones_ButtonClick(ByVal Button As MSComctlLib.Button)
  AccionTool Button.Index
End Sub
Private Function Validacion() As Boolean
  Validacion = True
  For II = 0 To 35
    If Campo(II).Tag = "1" And Campo(II).Text = "" Then Campo(II).Text = 0
  Next
  If Campo(24).Text <> "" Then
    If Campo(2).Text <> "" Then
      If Campo(5).Text <> "" Then
        If Campo(6).Text <> "" Then
          If Val(Campo(8).Text) <> 0 Then
            If Val(Campo(30).Text) <> 0 Then
              If Val(Campo(15).Text) <> 0 Then
                If Val(Campo(17).Text) <> 0 Then
                  If CboTpServicio.ListIndex = -1 Then
                    Validacion = False: MsgTit "Debe elegir algun tipo de servicio permitido por el cliente": CboTpServicio.SetFocus
                  Else
                    If CboTipo.ListIndex = 0 And ManejaCobroCorriente = 0 Then
                      Validacion = False: MsgTit "El cliente no maneja cobro corriente": CboTipo.SetFocus
                    Else
                      If CboTipo.ListIndex = 1 And ManejaCobroContado = 0 Then
                        Validacion = False: MsgTit "El cliente no maneja cobro contado": CboTipo.SetFocus
                      Else
                        If CboTipo.ListIndex = 2 And ManejaCobroDestino = 0 Then
                          Validacion = False: MsgTit "El cliente no maneja cobro destino": CboTipo.SetFocus
                        Else
                          Validacion = True
                        End If
                      End If
                    End If
                  End If
                Else
                  Validacion = False: MsgTit "Los kilos a facturar no pueden ser 0": Campo(16).SetFocus
                End If
              Else
                Validacion = False: MsgTit "La guia no se puede hacer con 0 unidades": Campo(15).SetFocus
              End If
            Else
              Validacion = False: MsgTit "La guia debe tener una ciudad de origen": Campo(30).SetFocus
            End If
          Else
            Validacion = False: MsgTit "La guia debe tener una ciudad de destino": Campo(8).SetFocus
          End If
        Else
          Validacion = False: MsgTit "El destinatario debe tener una direccion": Campo(6).SetFocus
        End If
      Else
        Validacion = False: MsgTit "La guia debe tener un destinatario": Campo(5).SetFocus
      End If
    Else
        Validacion = False: MsgTit "La guia debe tener un remitente": Campo(2).SetFocus
    End If
  Else
      Validacion = False: MsgTit "La guia debe tener una cuenta o tercero para el cobro": Campo(24).SetFocus
  End If
End Function

Sub Liquidacion()
  If NegociacionInactiva = True Or ListaPreciosVencida = True Then
    MsgBox "La negociacion de este cliente esta inactiva o la lista de precios esta vencida, debe activar la negociacion o cambiarle la fecha de vencimiento a la lista para poder liquidar la guia", vbCritical
  Else
    If PermiteRecaudo = False And Val(Campo(26).Text) <> 0 Then
        MsgBox "La negociacion de este cliente no permite tener recaudo, active esta opcion en clientes/negociacion", vbCritical
        Campo(26).Text = 0
    End If
      If Val(Campo(8).Text) <> 0 And Val(Campo(30).Text) <> 0 Then
        If Val(Campo(24).Text) <> 0 Then
          If Val(Campo(3).Text) <> 0 Then
            Load FrmLiquidacion
            If liquidado = True Then
              PasarDatosLiquidacion
            End If
              If Val(Campo(12).Text) <> 0 Then
                FrmLiquidacion.TxtDeclarado = ValNum(Campo(12))
              End If
              FrmLiquidacion.Show 1
              Campo(26).SetFocus
              liquidado = True
          Else
            MsgBox "Debe especificar una negociacion", vbCritical, "Falta negociacion"
            Campo(3).SetFocus
          End If
        Else
          MsgBox "Debe especificar una tercero", vbCritical, "Falta tercero"
          Campo(24).SetFocus
        End If
      Else
        MsgBox "Debe especificar una ciudad de destino y origen validos", vbCritical, "Falta ciudad destino"
        Campo(8).SetFocus
      End If
  End If
End Sub
Sub PasarDatosLiquidacion()
  FrmLiquidacion.TxtUnidades = ValNum(Campo(15))
  FrmLiquidacion.TxtKFacturar = ValNum(Campo(17))
  FrmLiquidacion.TxtKReales = ValNum(Campo(16))
  FrmLiquidacion.TxtKVolumen = ValNum(Campo(18))
  FrmLiquidacion.TxtVrManejo = ValNum(Campo(14))
  FrmLiquidacion.TxtVrFlete = ValNum(Campo(13))
  FrmLiquidacion.TxtDeclarado = ValNum(Campo(12))
    For II = 1 To 5
    If MProductos(II).IdProducto = 0 Then Exit For
    Set Item = FrmLiquidacion.LstTem.ListItems.Add(, , MProductos(II).Lote)
      Item.SubItems(1) = MProductos(II).IdProducto
      Item.SubItems(2) = MProductos(II).NmProducto
      Item.SubItems(3) = MProductos(II).IdEmpaque
      Item.SubItems(4) = MProductos(II).NmEmpaque
      Item.SubItems(5) = MProductos(II).Ancho
      Item.SubItems(6) = MProductos(II).Largo
      Item.SubItems(7) = MProductos(II).Alto
      Item.SubItems(8) = MProductos(II).Cantidad
      Item.SubItems(9) = MProductos(II).kilosReales
      Item.SubItems(10) = MProductos(II).KilosVol
      Item.SubItems(11) = MProductos(II).KilosFacturados
      Item.SubItems(12) = MProductos(II).VrFlete
  Next
End Sub
Sub LlenarGrillaProductos()
  Erase MProductos
    AbrirRecorset rstUniversal, "SELECT MvtoGuias.*, Productos.NmProducto AS Producto, Empaques.NmEmpaque AS Empaque FROM Empaques INNER JOIN MvtoGuias ON Empaques.IdEmpaque = MvtoGuias.IdEmpaque INNER JOIN Productos ON MvtoGuias.IdProducto = Productos.IdProducto where Guia=" & Val(Campo(0).Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    For II = 1 To rstUniversal.RecordCount
      MProductos(II).Lote = rstUniversal.Fields("Lote") & ""
      MProductos(II).IdProducto = rstUniversal.Fields("IdProducto")
      MProductos(II).NmProducto = rstUniversal.Fields("Producto")
      MProductos(II).IdEmpaque = rstUniversal.Fields("IdEmpaque")
      MProductos(II).NmEmpaque = rstUniversal.Fields("Empaque")
      MProductos(II).Ancho = rstUniversal.Fields("ancho")
      MProductos(II).Largo = rstUniversal.Fields("largo")
      MProductos(II).Alto = rstUniversal.Fields("altura")
      MProductos(II).Cantidad = rstUniversal.Fields("cant")
      MProductos(II).kilosReales = rstUniversal.Fields("KilosReal")
      MProductos(II).KilosVol = rstUniversal.Fields("KilosVol")
      MProductos(II).KilosFacturados = rstUniversal.Fields("KilosFacturados")
      MProductos(II).VrFlete = rstUniversal.Fields("VlrFlete")
      rstUniversal.MoveNext
    Next
    CerrarRecorset rstUniversal
End Sub
Sub VaciarGrilla()
  AbrirRecorset rstUniversal, "Delete from MvtoGuias where Guia=" & Val(Campo(0)), CnnPrincipal, adOpenDynamic, adLockOptimistic
  IniProg 1, 6
    For II = 1 To 6
      If MProductos(II).IdProducto = 0 Then Exit For
      Prog (II)
      AbrirRecorset rstUniversal, "INSERT INTO MvtoGuias (Guia, IdProducto, IdEmpaque, Largo, Ancho, Altura, KilosReal, KilosVol, Kilosfacturados, Cant, VlrFlete, Lote) VALUES (" & Val(Campo(0)) & "," & MProductos(II).IdProducto & "," & MProductos(II).IdEmpaque & "," & MProductos(II).Largo & "," & MProductos(II).Ancho & "," & MProductos(II).Largo & "," & MProductos(II).kilosReales & "," & MProductos(II).KilosVol & "," & MProductos(II).KilosFacturados & "," & MProductos(II).Cantidad & "," & Val(MProductos(II).VrFlete) & ",'" & MProductos(II).Lote & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
    Next
  Erase MProductos
  FinProg
  CerrarRecorset rstUniversal
End Sub


Private Sub ToolRemisiones_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  Select Case ButtonMenu.Tag
    Case "Buscar1"
      BuscarPorDocumento
    
    Case "Accion1" 'Quitar estado de impreso
      AbrirRecorset rstUniversal, "SELECT Guia, GuiFac FROM guias WHERE Guia = " & Val(Campo(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.RecordCount > 0 Then
        If ComprobarEstado(Val(Campo(0).Text)) = "I" Then
          If ComprobarEstadoSel(Val(Campo(0).Text), 0) = False And ComprobarEstadoSel(Val(Campo(0).Text), 5) = False Then
            If Val(rstUniversal!GuiFac) = 0 Then
              AbrirRecorset rstUniversal, "Update Guias set Estado='D' where Guia=" & Val(Campo(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
              MsgBox "Se ha actualizado el estado de la guia con exito", vbInformation
              Campo(19).Text = "D"
              AccionTool 17
            Else
              MsgBox "No se le puede quitar el estado de impreso a una guia factura"
            End If
          Else
            MsgBox "La guia no puede estar despachada, anulada o facturada para poder quitarle el estado de impreso", vbCritical
          End If
        Else
          MsgBox "la guia no esta impresa", vbCritical
        End If
      End If
      
    Case "Accion2" 'Cambiar numero de guia
      If ComprobarEstado(Val(Campo(0).Text)) = "D" Then
        If ComprobarEstadoSel(Val(Campo(0).Text), 0) = False Then
          If Principal.ToolConsultas1.AbrirDevDatos("Digite el numero de la guia", "Digite el numero para la guia que esta digitando", 3, 0) = True Then
            If ComprobarExGuia(Principal.ToolConsultas1.DatLo, Val(LblGuiaTipo.Tag)) = False Then
              AbrirRecorset rstUniversal, "Update Guias set Estado='D', Guia=" & Principal.ToolConsultas1.DatLo & " where Guia=" & Val(Campo(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
              AbrirRecorset rstUniversal, "Update recibos set GuiaRecibo=" & Principal.ToolConsultas1.DatLo & " where GuiaRecibo=" & Val(Campo(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
              
              MsgBox "Se ha actualizado el estado de la guia con exito", vbInformation
            Else
              MsgBox "La guia ya existe", vbCritical
            End If
          End If
        Else
          MsgBox "La guia no puede estar despachada, anulada o facturada para poder quitarle el estado de impreso", vbCritical
        End If
      Else
        MsgBox "La guia debe estar en estado digitada", vbCritical
      End If
      
    Case "Accion3" 'Reimprimir guia
      If ComprobarEstado(Val(Campo(0))) = "I" Then
        If GuiaFormato = True Then
          Mostrar_Reporte CnnPrincipal, 15, "Select*from sql_im_impguia where Guia=" & Val(Campo(0).Text), "", 2
          AccionTool 17
        Else
          FufuLo = SelectForm("Rem Cuartas", Me.hwnd)
          If ImprimirGuia(Campo(0).Text) = True Then
            AccionTool 17
            MsgBox "Se re-imprimio la guia con exito"
          End If
          establecerPapel
        End If
      Else
        MsgBox "La guia no esta en estado impreso", vbCritical
      End If
  End Select
    
End Sub

Private Sub CargarNegociacion()
  Dim rstListaPrecios As New ADODB.Recordset
  rstListaPrecios.CursorLocation = adUseClient

  AbrirRecorset rstUniversal, "SELECT negociaciones.* FROM negociaciones WHERE Id=" & Val(Campo(3)), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = False Then
      douPorcentajeManejo = rstUniversal!PorManejo
      douMinimoManejoUnidad = rstUniversal!MinUniManejo
      douMinimoManejoDespacho = rstUniversal!MinDesManejo
      douKilosMinimos = rstUniversal!Minimos
      douDctoKilo = rstUniversal!DctoK
      boolNoAplicarDctoReexpediciones = DevCheck(rstUniversal!NoAplicarDctoReexpediciones)
      LblConsulta(0) = rstUniversal!NmNegociacion & ""
      TpServicios(1) = rstUniversal!ManPaqueteo
      TpServicios(2) = rstUniversal!ManSemiMasivo
      TpServicios(3) = rstUniversal!ManMasivo
      TpServicios(4) = rstUniversal!ManLocal
      TpServicios(5) = rstUniversal!ManEncomiendas
      TpServicios(6) = 1
      TpServicios(7) = 1
      PermiteRecaudo = rstUniversal!PermiteRecaudo
      boolRedondearFlete = rstUniversal!RedondearFlete
      NegociacionInactiva = rstUniversal!Inactivo

      CPorte = rstUniversal!CartaPorte
      If CPorte <> 2 Then ChKCPorte.value = CPorte
      AbrirRecorset rstListaPrecios, "Select IdListaPrecios, FhVencimiento from listasprecios where IdListaPrecios=" & Val(rstUniversal.Fields("ListaPrecios")) & "", CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstListaPrecios.RecordCount > 0 Then
        If rstListaPrecios.Fields("FhVencimiento") <= Date Then
          ListaPreciosVencida = True
        Else
          ListaPreciosVencida = False
        End If
        intIdListaPrecios = rstListaPrecios!IdListaPrecios & ""
      End If
      CerrarRecorset rstListaPrecios
  Else
    LblConsulta(0) = "": Campo(3).Text = "0"
  End If
  CerrarRecorset rstUniversal
End Sub
Private Sub iniciarVariablesLiquidacion()
  douPorcentajeManejo = 0
  douMinimoManejoUnidad = 0
  douMinimoManejoDespacho = 0
  intIdListaPrecios = 0
  douVrKilo = 0
  douKilosMinimos = 0
  douDctoKilo = 0
  boolNoAplicarDctoReexpediciones = 0
  boolRedondearFlete = 0
End Sub

Private Function BuscarPorDocumento()
  If Principal.ToolConsultas1.AbrirDevDatos("Numero de documento", "Digite el numero del documento", 2, 0) = True Then
    AbrirRecorset rstUniversal, strSqlGuias & " WHERE DocCliente='" & Principal.ToolConsultas1.DatSt & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      Asignar rstUniversal
    Else
      MsgBox "No se encontraron guias con este documento", vbCritical
    End If
    CerrarRecorset rstUniversal
  End If
End Function

Private Sub LiquidarKilosFacturar()
  If Val(Campo(16).Text) > Val(Campo(17).Text) Then
    Campo(17).Text = Val(Campo(16).Text)
  End If
  If Val(Campo(18).Text) > Val(Campo(17).Text) Then
    Campo(17).Text = Val(Campo(18).Text)
  End If
End Sub
