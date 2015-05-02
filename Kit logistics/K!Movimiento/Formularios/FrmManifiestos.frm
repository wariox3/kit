VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManifiestos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Despachos..."
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10740
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdInformacionDespachoLiquidar 
      Caption         =   "Liquidar"
      Height          =   255
      Left            =   1800
      TabIndex        =   95
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Frame FraOtros 
      Caption         =   "Otros"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   3240
      TabIndex        =   90
      Top             =   3840
      Width           =   3855
      Begin VB.TextBox TxtPagoDescargue 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox TxtPagoCargue 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Descargue:"
         Height          =   195
         Left            =   120
         TabIndex        =   92
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cargue:"
         Height          =   195
         Left            =   360
         TabIndex        =   91
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.TextBox TxtCampos 
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
      Left            =   9240
      TabIndex        =   89
      Tag             =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox TxtCampos 
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
      Left            =   9600
      TabIndex        =   88
      Tag             =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox TxtCampos 
      Height          =   285
      Index           =   34
      Left            =   9960
      TabIndex        =   87
      Tag             =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton CmdDespachosPendientes 
      Caption         =   "Ver despachos pendientes"
      Height          =   255
      Left            =   3600
      TabIndex        =   81
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   2160
      TabIndex        =   67
      Top             =   1320
      Width           =   2895
      Begin VB.TextBox TxtTotalCE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   84
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   33
         Left            =   1320
         TabIndex        =   82
         Tag             =   "1"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   18
         Left            =   1320
         TabIndex        =   70
         Tag             =   "1"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TxtCampos 
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
         Left            =   1320
         TabIndex        =   69
         Tag             =   "1"
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Manejo CE:"
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   94
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Flete CE:"
         Height          =   195
         Index           =   7
         Left            =   480
         TabIndex        =   93
         Top             =   960
         Width           =   645
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Recaudos:"
         Height          =   195
         Index           =   10
         Left            =   405
         TabIndex        =   83
         Top             =   240
         Width           =   780
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Total CE:"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   68
         Top             =   600
         Width           =   660
      End
   End
   Begin VB.Frame FraCit 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   5160
      TabIndex        =   44
      Top             =   1320
      Width           =   5415
      Begin VB.TextBox TxtCampos 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   48
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox CboTipo 
         Height          =   315
         ItemData        =   "FrmManifiestos.frx":0000
         Left            =   1560
         List            =   "FrmManifiestos.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox TxtCampos 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   840
         TabIndex        =   2
         Tag             =   "1"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox TxtNmRuta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   51
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox TxtCampos 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   840
         TabIndex        =   1
         Tag             =   "1"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox TxtNmCiudad 
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   1560
         TabIndex        =   46
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox TxtNmCiudad 
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   1560
         TabIndex        =   45
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox TxtCampos 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   840
         TabIndex        =   0
         Tag             =   "1"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   71
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   165
         TabIndex        =   50
         Top             =   600
         Width           =   585
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Index           =   24
         Left            =   360
         TabIndex        =   49
         Top             =   960
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Origen:"
         Height          =   195
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Frame FraInfo 
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   41
      Top             =   615
      Width           =   10455
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   35
         Left            =   5760
         TabIndex        =   86
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   10
         Left            =   8160
         TabIndex        =   72
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   31
         Left            =   4320
         TabIndex        =   58
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   57
         Tag             =   "1"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   56
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Man elect:"
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
         Index           =   6
         Left            =   4800
         TabIndex        =   85
         Top             =   240
         Width           =   915
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "CO:"
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
         Index           =   2
         Left            =   3960
         TabIndex        =   55
         Top             =   240
         Width           =   330
      End
      Begin VB.Label LblEstado 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   8640
         TabIndex        =   53
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label LblUniversal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
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
         Index           =   0
         Left            =   7440
         TabIndex        =   52
         Top             =   240
         Width           =   660
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Manifiesto:"
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
         Index           =   19
         Left            =   2040
         TabIndex        =   43
         Top             =   240
         Width           =   945
      End
      Begin VB.Label LblUniversal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Despacho:"
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
         Index           =   51
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.Frame FraResumen 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      TabIndex        =   38
      Top             =   1320
      Width           =   1935
      Begin VB.TextBox TxtCampos 
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
         Index           =   14
         Left            =   960
         TabIndex        =   66
         Tag             =   "1"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   11
         Left            =   960
         TabIndex        =   63
         Tag             =   "1"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   13
         Left            =   960
         TabIndex        =   62
         Tag             =   "1"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   12
         Left            =   960
         TabIndex        =   61
         Tag             =   "1"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "kilos Vol:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   65
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Remesas:"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   64
         Top             =   240
         Width           =   705
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "kilos Real:"
         Height          =   195
         Index           =   49
         Left            =   135
         TabIndex        =   40
         Top             =   960
         Width           =   735
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Unidades:"
         Height          =   195
         Index           =   47
         Left            =   150
         TabIndex        =   39
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.Frame FraTotales 
      Caption         =   "Resumen"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   6720
      TabIndex        =   33
      Top             =   5280
      Width           =   3855
      Begin VB.TextBox TxtCampos 
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
         Index           =   30
         Left            =   1800
         TabIndex        =   76
         Tag             =   "1"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   29
         Left            =   1800
         TabIndex        =   75
         Tag             =   "1"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   28
         Left            =   1800
         TabIndex        =   74
         Tag             =   "1"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   25
         Left            =   1800
         TabIndex        =   73
         Tag             =   "1"
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Total Viaje:"
         Height          =   195
         Index           =   23
         Left            =   825
         TabIndex        =   37
         Top             =   960
         Width           =   795
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Manifiesto:"
         Height          =   195
         Index           =   46
         Left            =   405
         TabIndex        =   36
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Descuentos Varios:"
         Height          =   195
         Index           =   31
         Left            =   240
         TabIndex        =   35
         Top             =   1320
         Width           =   1380
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Flete adicional:"
         Height          =   195
         Index           =   41
         Left            =   555
         TabIndex        =   34
         Top             =   600
         Width           =   1065
      End
   End
   Begin VB.Frame FraDescuentos 
      Caption         =   "Otros descuentos"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   3240
      TabIndex        =   28
      Top             =   5280
      Width           =   3375
      Begin VB.TextBox TxtCampos 
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
         Index           =   24
         Left            =   1200
         TabIndex        =   14
         Tag             =   "1"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   23
         Left            =   1200
         TabIndex        =   13
         Tag             =   "1"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   22
         Left            =   1200
         TabIndex        =   12
         Tag             =   "1"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   21
         Left            =   1200
         TabIndex        =   11
         Tag             =   "1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Cargue:"
         Height          =   195
         Index           =   34
         Left            =   510
         TabIndex        =   32
         Top             =   960
         Width           =   555
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Estampilla:"
         Height          =   195
         Index           =   35
         Left            =   315
         TabIndex        =   31
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Seguridad:"
         Height          =   195
         Index           =   33
         Left            =   300
         TabIndex        =   30
         Top             =   600
         Width           =   765
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Papeleria:"
         Height          =   195
         Index           =   32
         Left            =   360
         TabIndex        =   29
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame FraFlete 
      Caption         =   "Flete y anticipo"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      TabIndex        =   25
      Top             =   5280
      Width           =   3015
      Begin VB.TextBox TxtCampos 
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
         Enabled         =   0   'False
         Height          =   285
         Index           =   26
         Left            =   1080
         TabIndex        =   78
         Tag             =   "1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox TxtCampos 
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
         Enabled         =   0   'False
         Height          =   285
         Index           =   27
         Left            =   1080
         TabIndex        =   77
         Tag             =   "1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   20
         Left            =   1080
         TabIndex        =   10
         Tag             =   "1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   19
         Left            =   1080
         TabIndex        =   9
         Tag             =   "1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Ret. Fuente:"
         Height          =   195
         Index           =   29
         Left            =   120
         TabIndex        =   80
         Top             =   960
         Width           =   885
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Ind. Com:"
         Height          =   195
         Index           =   30
         Left            =   240
         TabIndex        =   79
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "Valor Flete:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   795
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Anticipo:"
         Height          =   195
         Index           =   36
         Left            =   300
         TabIndex        =   26
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame FraDatos 
      Caption         =   "Observaciones"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   7200
      TabIndex        =   24
      Top             =   3840
      Width           =   3375
      Begin VB.TextBox TxtCampos 
         Height          =   975
         Index           =   32
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton CmdCargarRemisiones 
      Caption         =   "Cargar guias al despacho"
      Height          =   255
      Left            =   6720
      TabIndex        =   16
      Top             =   7080
      Width           =   3855
   End
   Begin VB.CommandButton CmdVerRemesas 
      Caption         =   "Ver Remesas"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Frame FraVeCon 
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   10455
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   5
         Left            =   960
         MaxLength       =   6
         TabIndex        =   4
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   6
         Left            =   3120
         TabIndex        =   5
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vehiculo:"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   180
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Conductor:"
         Height          =   195
         Left            =   2280
         TabIndex        =   22
         Top             =   180
         Width           =   780
      End
      Begin VB.Label LblNmConductor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4440
         TabIndex        =   21
         Top             =   180
         Width           =   5895
      End
   End
   Begin VB.Frame FraEnc 
      Caption         =   "Fechas"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   3015
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   4
         Left            =   600
         TabIndex        =   60
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   3
         Left            =   600
         TabIndex        =   59
         Top             =   255
         Width           =   2175
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Exp:"
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   19
         Top             =   285
         Width           =   315
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Cump:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   450
      End
   End
   Begin MSComctlLib.Toolbar ToolDespachos 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   10740
      _ExtentX        =   18944
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
            Object.ToolTipText     =   "Buscar [Inicio] por despacho"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "B1"
                  Text            =   "Buscar por manifiesto"
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "IMP1"
                  Text            =   "Impresora punto"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "IMP2"
                  Text            =   "Formato encabezado"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "IMP3"
                  Text            =   "Formato detalle"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Car"
            Object.ToolTipText     =   "Carga informacion adicional [Pausa]"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Otro"
            Style           =   5
            Object.Width           =   1e-4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "O1"
                  Text            =   "Imprir orden de despacho"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "O2"
                  Text            =   "Relacion de entrega"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "O8"
                  Text            =   "Relacion contraentregas y recaudo"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "O6"
                  Text            =   "Generar monitoreo"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "O15"
                  Text            =   "Generar manifiesto interno"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmManifiestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstDespachos As New ADODB.Recordset
Dim Editando As Boolean
Dim RteFte As Currency
Dim VrMayorA As Currency
Dim IndCom As Currency
Dim DctoPapeleria As Currency
Dim Estampilla As Currency
Dim strSqlDespachos As String


Private Sub CboTipo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub



Private Sub CmdCargarRemisiones_Click()
  If TxtCampos(10) = "D" Then
    FufuLo = TxtCampos(1)
    FrmLlenarDespacho.Show 1
    AbrirRecorset rstUniversal, "Update Despachos set Remesas=" & Val(Format(TxtCampos(11).Text, "0;(0)")) & ", Unidades=" & Val(Format(TxtCampos(12).Text, "0;(0)")) & ", KilosReales=" & Val(Format(TxtCampos(13).Text, "0;(0)")) & ", KilosVol=" & Val(Format(TxtCampos(14).Text, "0;(0)")) & ", FleteCobra=" & Val(Format(TxtCampos(15).Text, "0;(0)")) & ", VrDeclaradoTotal=" & Val(Format(TxtCampos(34).Text, "0;(0)")) & ", ManejoCobra=" & Val(Format(TxtCampos(16).Text, "0;(0)")) & ", TotalCE=" & Val(Format(TxtTotalCE.Text, "0;(0)")) & ", FleteCE=" & Val(Format(TxtCampos(17).Text, "0;(0)")) & ", ManejoCE=" & Val(Format(TxtCampos(18).Text, "0;(0)")) & ",  TRecaudo=" & Val(Format(TxtCampos(33).Text, "0;(0)")) & " where OrdDespacho=" & TxtCampos(1), CnnPrincipal, adOpenDynamic, adLockOptimistic
    AccionTool 17
    'TxtTotalCE = Format(Val(Format(TxtCampos(17).Text, "0;(0)")) + Val(Format(TxtCampos(18).Text, "0;(0)")) + Val(Format(TxtCampos(33).Text, "0;(0)")), "#,##0;(#,##0)")
  Else
    MsgBox "Solo se le pueden agregar guias a los despachos en estado [DIGITADO]", vbCritical
  End If
End Sub

Private Sub CmdDespachosPendientes_Click()
  FrmDespachosPendientes.Show 1
End Sub





Private Sub CmdInformacionDespachoLiquidar_Click()
  Cobros Val(TxtCampos(1).Text)
  FufuLo = Val(TxtCampos(1).Text)
  FrmInformacionLiquidarDespacho.Show 1
  'Cobros Val(TxtCampos(1).Text)
End Sub

Private Sub CmdVerRemesas_Click()
  If CpPermisoEspecial(11, CodUsuarioActivo, CnnPrincipal) = True Then
    FufuLo = Val(TxtCampos(1).Text)
    FrmVerGuiasDespacho.Show 1
  Else
    MsgBox "No tiene permisos para ver esta informacion"
  End If

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    LaTecla (KeyCode), ToolDespachos
End Sub

Private Sub Form_Load()
  IconosTool ToolDespachos, Principal.IgListTool
  strSqlDespachos = "SELECT despachos.*, " & _
                              "CiudadOrigen.NmCiudad AS NmCiudadOrigen, " & _
                              "CiudadDestino.NmCiudad AS NmCiudadDestino, " & _
                              "rutas.NmRuta " & _
                              "FROM despachos " & _
                              "LEFT JOIN ciudades AS CiudadOrigen ON despachos.IdCiudadOrigen = CiudadOrigen.IdCiudad " & _
                              "LEFT JOIN ciudades AS CiudadDestino ON despachos.IdCiudadDestino = CiudadDestino.IdCiudad " & _
                              "LEFT JOIN rutas ON despachos.IdRuta = rutas.IdRuta "
                              
  AbrirRecorset rstDespachos, strSqlDespachos & " Order by OrdDespacho Desc Limit 50", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Formatos rstDespachos
  Asignar rstDespachos
  
  AbrirRecorset rstUniversal, "Select RteFte, RteFteMayor, IndCom, DctoPapeleria, Estampilla from ParametrizacionLiquidaciones", CnnPrincipal, adOpenDynamic, adLockOptimistic
    RteFte = rstUniversal!RteFte
    VrMayorA = rstUniversal!RteFteMayor
    IndCom = rstUniversal!IndCom
    DctoPapeleria = rstUniversal!DctoPapeleria
    Estampilla = rstUniversal!Estampilla
  CerrarRecorset rstUniversal
End Sub
Private Sub Formatos(rstForma As ADODB.Recordset)
  For II = 0 To 35
    Set rstForma.Fields(II).DataFormat = TxtCampos(II).DataFormat
  Next
End Sub
Private Sub Asignar(rstAsignar As ADODB.Recordset)
  For II = 0 To 35
    TxtCampos(II).Text = rstAsignar.Fields(II)
  Next
  CboTipo.ListIndex = Val(TxtCampos(2).Text) - 1
  TxtNmCiudad(8).Text = rstAsignar!NmCiudadOrigen & ""
  TxtNmCiudad(9).Text = rstAsignar!NmCiudadDestino & ""
  TxtNmRuta.Text = rstAsignar!NmRuta & ""
  LblNmConductor.Caption = rstAsignar.Fields("NmConductor") & ""
  TxtPagoCargue.Text = rstAsignar!PagoCargue & ""
  TxtPagoDescargue.Text = rstAsignar!PagoDescargue & ""
  TxtTotalCE = rstAsignar!TotalCE
End Sub
Private Sub limpiar()
  For II = 0 To 35
    TxtCampos(II).Text = ""
  Next
  CboTipo.ListIndex = 0
  TxtNmCiudad(8).Text = ""
  TxtNmCiudad(9).Text = ""
  TxtNmRuta.Text = ""
  LblNmConductor.Caption = ""
  TxtPagoCargue.Text = ""
  TxtPagoDescargue.Text = ""
End Sub
Private Sub Desbloquear()
  FraDatos.Enabled = True
  FraOtros.Enabled = True
  FraCit.Enabled = True
  FraVeCon.Enabled = True
  FraFlete.Enabled = True
  FraDescuentos.Enabled = True
  BotTool 3, 17, ToolDespachos, True
  CmdCargarRemisiones.Enabled = False
  CmdVerRemesas.Enabled = False

End Sub
Private Sub Bloquear()
  FraDatos.Enabled = False
  FraOtros.Enabled = False
  FraCit.Enabled = False
  FraVeCon.Enabled = False
  FraFlete.Enabled = False
  FraDescuentos.Enabled = False
  BotTool 3, 17, ToolDespachos, False
  CmdCargarRemisiones.Enabled = True
  CmdVerRemesas.Enabled = True

End Sub

Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
      Desbloquear
      limpiar
      TxtCampos(31) = Coperaciones
      TxtCampos(3) = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss AM/PM")
      TxtCampos(8).SetFocus
      TxtCampos(10) = "D"
      AbrirRecorset rstUniversal, "SELECT Nombre FROM informacionempresa WHERE Id = 1", CnnPrincipal, adOpenDynamic, adLockOptimistic
      TxtPagoCargue.Text = rstUniversal.Fields("Nombre") & ""
      TxtPagoDescargue.Text = rstUniversal.Fields("Nombre") & ""
      CerrarRecorset rstUniversal
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            Bloquear
            AbrirRecorset rstUniversal, "Update Despachos set Tipo=" & CboTipo.ListIndex + 1 & ", IdVehiculo='" & TxtCampos(5) & "', IdConductor='" & TxtCampos(6) & "', IdRuta=" & TxtCampos(7) & ", IdCiudadOrigen=" & TxtCampos(8) & ", IdCiudadDestino=" & TxtCampos(9) & ", VrFlete=" & Val(Format(TxtCampos(19).Text, "0;(0)")) & ", VrAnticipo=" & Val(Format(TxtCampos(20).Text, "0;(0)")) & ", VrDctoPapeleria=" & Val(Format(TxtCampos(21).Text, "0;(0)")) & ", VrDctoSeguridad=" & Val(Format(TxtCampos(22).Text, "0;(0)")) & ", " & _
            " VrDctoCargue=" & Val(Format(TxtCampos(23).Text, "0;(0)")) & ", VrDctoEstampilla=" & Val(Format(TxtCampos(24).Text, "0;(0)")) & ", VrFleteAdicional=" & Val(Format(TxtCampos(25).Text, "0;(0)")) & ", VrDctoIndCom=" & Val(Format(TxtCampos(26).Text, "0;(0)")) & ", VrDctoRteFte=" & Val(Format(TxtCampos(27).Text, "0;(0)")) & ", VrOtrosDctos=" & Val(Format(TxtCampos(28).Text, "0;(0)")) & ", SaldoDesp=" & Val(Format(TxtCampos(29).Text, "0;(0)")) & ", TotalViaje=" & Val(Format(TxtCampos(30).Text, "0;(0)")) & ", TRecaudo=" & Val(Format(TxtCampos(33).Text, "0;(0)")) & " , VrDeclaradoTotal=" & Val(Format(TxtCampos(34).Text, "0;(0)")) & ", Observaciones='" & TxtCampos(32) & "', NmConductor='" & LblNmConductor.Caption & _
            "', PagoCargue = '" & TxtPagoCargue.Text & "' , PagoDescargue = '" & TxtPagoDescargue.Text & "' where OrdDespacho=" & TxtCampos(1), CnnPrincipal, adOpenDynamic, adLockOptimistic
            Editando = False
            AccionTool 17
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          
          TxtCampos(1) = SacarConsecutivo("Despachos", CnnPrincipal)
          AbrirRecorset rstUniversal, "INSERT INTO Despachos (IdManifiesto, OrdDespacho, Tipo, FhExpedicion, FhCumplidos, IdVehiculo, IdConductor, IdRuta, IdCiudadOrigen, IdCiudadDestino, Estado, Remesas, Unidades, KilosReales, KilosVol, FleteCobra, ManejoCobra, FleteCE, ManejoCE, VrFlete, VrAnticipo, VrDctoPapeleria, VrDctoSeguridad, VrDctoCargue, VrDctoEstampilla, VrFleteAdicional, VrDctoIndCom, VrDctoRteFte, VrOtrosDctos, SaldoDesp, TotalViaje, CO, Observaciones, TRecaudo, VrDeclaradoTotal, Cerrado, Liquidado, IdUsuario, NmConductor, IdEmpresa, PagoCargue, PagoDescargue)" & _
          " VALUES (" & TxtCampos(0) & ", " & TxtCampos(1) & ", " & CboTipo.ListIndex + 1 & ", now(), now(), '" & UCase(TxtCampos(5).Text) & "', '" & TxtCampos(6).Text & "', " & TxtCampos(7) & ", " & TxtCampos(8) & ", " & TxtCampos(9) & ", '" & TxtCampos(10) & "', " & Val(Format(TxtCampos(11).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(12).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(13).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(14).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(15).Text, "0;(0)")) & _
          ", " & Val(Format(TxtCampos(16).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(17).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(18).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(19).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(20).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(21).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(22).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(23).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(24).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(25).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(26).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(27).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(28).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(29).Text, "0;-0")) & ", " & _
          " " & Val(Format(TxtCampos(30).Text, "0;-0")) & ", " & TxtCampos(31) & ", '" & TxtCampos(32) & "', " & Val(Format(TxtCampos(33).Text, "0;-0")) & " , " & Val(Format(TxtCampos(34).Text, "0;(0)")) & ",0,0, " & CodUsuarioActivo & ",'" & LblNmConductor.Caption & "', " & CodEmpresaActiva & ", '" & TxtPagoCargue.Text & "', '" & TxtPagoDescargue.Text & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
          
          Bloquear
          TxtCampos(10) = "D"
        End If
      End If
    Case 5  'Editar
      If TxtCampos(10) = "D" Then
        Editando = True
        Desbloquear
        TxtCampos(8).SetFocus
      Else
        MsgBox "Solo se pueden editar los despachos digitados", vbCritical
      End If
    Case 6 'Eliminar
      If TxtCampos(10).Text = "D" Or TxtCampos(10).Text = "V" Then
        If MsgBox("Un despacho no se puede eliminar, unicamente anular ¿Desea anular el despacho?", vbYesNo + vbQuestion) = vbYes Then
          If MsgBox("¿Desea cargar estas guias en un temporal para agregar a un nuevo despacho?", vbQuestion + vbYesNo) = vbYes Then
            Dim rstInsertar As New ADODB.Recordset
            rstInsertar.CursorLocation = adUseClient
            AbrirRecorset rstUniversal, "TRUNCATE temp_guias_despacho_anulado", CnnPrincipal, adOpenDynamic, adLockOptimistic
            AbrirRecorset rstUniversal, "SELECT Guia FROM guias WHERE IdDespacho = " & Val(TxtCampos(1).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
            Do While rstUniversal.EOF = False
              AbrirRecorset rstInsertar, "INSERT INTO temp_guias_despacho_anulado (Guia) VALUES(" & rstUniversal.Fields("Guia") & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
              rstUniversal.MoveNext
            Loop
            CerrarRecorset rstUniversal
          End If
          AbrirRecorset rstUniversal, "Update Despachos set Estado='A' where OrdDespacho=" & Val(TxtCampos(1).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
          AbrirRecorset rstUniversal, "Update Guias set Estado='I', IdDespacho=null, Despachada=0 where IdDespacho=" & TxtCampos(1).Text, CnnPrincipal, adOpenDynamic, adLockOptimistic
          TxtCampos(10).Text = "A"
          MsgBox "El despacho se anulo con exito y las guias fueron liberadas para agregar a un nuevo despacho", vbExclamation, "Guia anulada con exito"
          AccionTool 17

        End If
      Else
        MsgBox "Solo se pueden anular los despachos digitados y viajando", vbCritical
      End If
    
    Case 7 'Cancelar
      If Editando = True Then
        AccionTool 4
      Else
        Asignar rstDespachos
        Bloquear
      End If
    Case 9  'Buscar
      If Principal.ToolConsultas1.AbrirDevDatos("Numero de despacho", "Digite el numero del despacho que desea buscar", 3, 0) = True Then
        AbrirRecorset rstUniversal, strSqlDespachos & " WHERE OrdDespacho=" & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          Formatos rstUniversal
          Asignar rstUniversal
        Else
          MsgBox "No se encontraron despachos con este numero", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If
    Case 11 'Primero
      UPrimero rstDespachos
      Asignar rstDespachos
    Case 12 'Anterior
      UAnterior rstDespachos
      Asignar rstDespachos
    Case 13 'Siguiente
      USiguiente rstDespachos
      Asignar rstDespachos
    Case 14 'Ultimo
      UUltimo rstDespachos
      Asignar rstDespachos
    Case 16 'Cerrar
      CerrarRecorset rstDespachos
      Unload Me
    Case 17 'Actualizar
      rstDespachos.Requery
      Formatos rstDespachos
    Case 18 'Imprimir
      
    Case 19 'Recargar
      If Val(TxtCampos(8).Text) <> 0 Then TxtNmCiudad(8).Text = DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & TxtCampos(8), "NmCiudad", CnnPrincipal)
      If Val(TxtCampos(9).Text) <> 0 Then TxtNmCiudad(9).Text = DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & TxtCampos(9), "NmCiudad", CnnPrincipal)
      If Val(TxtCampos(7).Text) <> 0 Then TxtNmRuta.Text = DevResBus("SELECT IdRuta, NmRuta From Rutas where IdRuta=" & TxtCampos(7), "NmRuta", CnnPrincipal)
      If TxtCampos(6).Text <> "" Then
        AbrirRecorset rstUniversal, "Select IdConductor, concat(Nombre, ' ', Apellido1, ' ', Apellido2) as NmConductor From Conductores where IdConductor='" & TxtCampos(6).Text & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          LblNmConductor.Caption = rstUniversal.Fields("NmConductor")
        End If
        CerrarRecorset rstUniversal
      End If
  End Select
End Sub

Private Function Validacion() As Boolean
  Validacion = True
  For II = 0 To 35
    If TxtCampos(II).Tag = "1" And TxtCampos(II).Text = "" Then
      TxtCampos(II) = 0
    End If
  Next
  If Val(TxtCampos(8).Text) <> 0 Then
    If Val(TxtCampos(9).Text) <> 0 Then
      If Val(TxtCampos(7).Text) <> 0 Then
        If TxtCampos(5).Text <> "" Then
          If TxtCampos(6).Text <> "" Then
            Validacion = True
          Else
            Validacion = False: MsgTit "El despacho debe tener un conductor": TxtCampos(6).SetFocus
          End If
        Else
          Validacion = False: MsgTit "Debe especificar un vehiculo": TxtCampos(5).SetFocus
        End If
      Else
        Validacion = False: MsgTit "El despacho debe tener una ruta": TxtCampos(7).SetFocus
      End If
    Else
      Validacion = False: MsgTit "El despacho debe tener un destino": TxtCampos(9).SetFocus
    End If
  Else
    Validacion = False: MsgTit "El despacho debe tener un origen": TxtCampos(8).SetFocus
  End If
End Function



Private Sub ToolDespachos_ButtonClick(ByVal Button As MSComctlLib.Button)
  AccionTool Button.Index
End Sub

Private Sub ToolDespachos_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  Select Case ButtonMenu.Tag
    Case "O1"
      Mostrar_Reporte CnnPrincipal, 4, "Select*from sql_im_OrdenDespacho where IdDespacho=" & Val(TxtCampos(1)), "", 2
    
    Case "O2"
      AbrirRecorset rstUniversal, "SELECT Estado FROM despachos WHERE OrdDespacho = " & Val(TxtCampos(1)), CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.RecordCount > 0 Then
        If rstUniversal.Fields("Estado") = "D" Then
          Dim rstActualizar As New ADODB.Recordset
          rstActualizar.CursorLocation = adUseClient
          AbrirRecorset rstActualizar, "UPDATE despachos SET Estado = 'V' WHERE OrdDespacho = " & Val(TxtCampos(1).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
          AbrirRecorset rstActualizar, "Update Guias set Estado='V' where IdDespacho=" & Val(TxtCampos(1).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
        End If
      End If
      CerrarRecorset rstUniversal
      Mostrar_Reporte CnnPrincipal, 6, "Select*from sql_im_rel_guias_desp where IdDespacho=" & Val(TxtCampos(1)), "", 2
      
    Case "O6"
      AbrirRecorset rstUniversal, "Insert into MonitoreoVehiculos (Orden, Tipo, Estado, Ok, FhHrSalida, Vehiculo, Destino, UltReporte, Frecuencia, EnNovedad, SinMonitoreo) Values (" & Val(TxtCampos(1)) & ", " & CboTipo.ListIndex + 1 & ", 'P', 0, '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "','" & TxtCampos(5) & "', '" & DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Val(TxtCampos(9)), "NmCiudad", CnnPrincipal) & "', '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', 0, 0, 0)", CnnPrincipal, adOpenDynamic, adLockOptimistic
      MsgBox "Se genero el monitoreo con exito", vbExclamation
      
    Case "O8"
      Mostrar_Reporte CnnPrincipal, 8, "Select*from sql_im_contraentregas where IdDespacho=" & Val(TxtCampos(1).Text), "Contraentregas y recaudos", 2
      
    Case "IMP2"
      Mostrar_Reporte CnnPrincipal, 16, "Select*from sql_im_imprimirmanifiesto where OrdDespacho=" & Val(TxtCampos(1)), "", 2
    Case "IMP3"
      Mostrar_Reporte CnnPrincipal, 29, "Select*from sql_im_imprimirmanifiestodetalle where OrdDespacho=" & Val(TxtCampos(1)), "", 2
    Case "IMP1"
      AbrirRecorset rstUniversal, "SELECT OrdDespacho, IdManifiesto FROM despachos WHERE IdManifiesto != 0 AND OrdDespacho = " & Val(TxtCampos(1).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.RecordCount > 0 Then
        CerrarRecorset rstUniversal
        ImprimirManifiesto (Val(TxtCampos(1).Text))
      Else
        MsgBox "Debe generar el manifiesto interno antes de imprimir", vbCritical, "Error al imprimir manifiesto"
      End If
      CerrarRecorset rstUniversal
    Case "O15"
      If TxtCampos(10).Text = "D" Then
        If Val(TxtCampos(8).Text) <> 0 Then
          If Val(TxtCampos(9).Text) <> 0 Then
            If Val(TxtCampos(7).Text) <> 0 Then
              If TxtCampos(5).Text <> "" Then
                If TxtCampos(6).Text <> "" Then
                  If Format(TxtCampos(19).Text, "###0.00") > 0 Then
                    If Val(TxtCampos(11).Text) > 0 Then
                      If CboTipo.ListIndex = 0 Then
                        If CpPermisoEspecial(13, CodUsuarioActivo, CnnPrincipal) = True Then
                          Dim rstVehiculo As New ADODB.Recordset
                          rstVehiculo.CursorLocation = adUseClient
                          AbrirRecorset rstVehiculo, "SELECT IdMarca, IdLinea FROM vehiculos WHERE IdPlaca = '" & TxtCampos(5).Text & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
                          If rstVehiculo.Fields("IdMarca") & "" <> "" And rstVehiculo.Fields("IdLinea") & "" <> "" Then
                            FufuLo = SacarConsecutivo("Manifiestos", CnnPrincipal)
                            AbrirRecorset rstUniversal, "Update Despachos set Estado='V', FhPagoSaldo='" & Format(Date + 8, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', IdManifiesto=" & FufuLo & " where OrdDespacho=" & TxtCampos(1).Text, CnnPrincipal, adOpenDynamic, adLockOptimistic
                            AbrirRecorset rstUniversal, "Update Guias set Estado='V' where IdDespacho=" & TxtCampos(1).Text, CnnPrincipal, adOpenDynamic, adLockOptimistic
                            AbrirRecorset rstUniversal, "Insert into monitoreovehiculos (Orden, Tipo, Estado, Ok, FhHrSalida, Vehiculo, Destino, UltReporte, Frecuencia, EnNovedad, SinMonitoreo) Values (" & Val(TxtCampos(1)) & ", " & CboTipo.ListIndex + 1 & ", 'P', 0, '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "','" & TxtCampos(5) & "', '" & DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Val(TxtCampos(9)), "NmCiudad", CnnPrincipal) & "', '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', 0, 0, 0)", CnnPrincipal, adOpenDynamic, adLockOptimistic
                            AbrirRecorset rstUniversal, "Insert into despachos_control_mt (OrdDespacho, ManifiestoInterno) Values (" & Val(TxtCampos(1)) & ", " & FufuLo & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
                            TxtCampos(0).Text = FufuLo
                            TxtCampos(10).Text = "V"
                            Cobros Val(TxtCampos(1).Text)
                            MsgBox "El manifiesto interno se genero con exito, se va a crear un registro para el monitoreo del vehiculo automaticamente", vbExclamation
                            AccionTool 17
                          Else
                            MsgBox "El vehiculo no tiene marca o linea por favor verifique"
                          End If
                          CerrarRecorset rstVehiculo
                        Else
                          MsgBox "El usuario no esta autorizado para generar manifiestos internos"
                        End If
                      Else
                        MsgBox "Solo se pueden imprimir los despachos de viaje", vbCritical
                      End If
                    Else
                      MsgBox "El manifiesto debe tener guias para poderse imprimir", vbCritical
                    End If
                  Else
                    MsgBox "El manifiesto debe tener un flete mayor a cero", vbCritical
                  End If
                Else
                  MsgBox "El despacho debe tener un conductor", vbCritical
                End If
              Else
                MsgBox "El despacho debe tener un vehiculo asignado", vbCritical
              End If
            Else
              MsgBox "El despacho no tiene ruta, no se puede imprimir", vbCritical
            End If
          Else
            MsgBox "El despacho no tiene ciudad de destino, no se puede imprimir", vbCritical
          End If
        Else
          MsgBox "El despacho no tiene ciudad de origen, no se puede imprimir", vbCritical
        End If
      Else
        MsgBox "Solo se pueden imprimir los despachos en estado [DIGITADO]", vbCritical
      End If
      Case "B1"
        If Principal.ToolConsultas1.AbrirDevDatos("Numero de manifiesto", "Digite el numero del manifiesto que desea buscar", 3, 0) = True Then
          AbrirRecorset rstUniversal, strSqlDespachos & " WHERE IdManifiesto=" & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
          If rstUniversal.EOF = False Then
            Formatos rstUniversal
            Asignar rstUniversal
          Else
            MsgBox "No se encontraron despachos con este numero de manifiesto", vbCritical
          End If
          CerrarRecorset rstUniversal
        End If
  End Select
End Sub

Private Sub TxtCampos_Change(Index As Integer)
  If Index = 10 Then LblEstado.Caption = DevEstadoDespacho(TxtCampos(10))
End Sub

Private Sub TxtCampos_GotFocus(Index As Integer)
  EnfocarT TxtCampos(Index)
  TxtCampos(Index).BackColor = &H80000001
  TxtCampos(Index).ForeColor = &HFFFFFF
End Sub

Private Sub TxtCampos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Select Case Index
      Case 8
        Principal.ToolConsultas1.AbrirDevConsulta 1, CnnPrincipal
        TxtCampos(8).Text = Principal.ToolConsultas1.DatLo
        
      Case 9
        FufuLo = 0
        FufuSt = TxtCampos(8).Tag
        'FrmSeleccionarDestino.Show 1
        'TxtCampos(9).Text = FufuLo
        
        Principal.ToolConsultas1.AbrirDevConsulta 1, CnnPrincipal
        TxtCampos(9).Text = Principal.ToolConsultas1.DatLo
        
      Case 7
        Principal.ToolConsultas1.AbrirConsultaGral "IdRuta", "NmRuta", "Rutas", CnnPrincipal
        TxtCampos(7).Text = Principal.ToolConsultas1.DatLo
      
      Case 5
        Principal.ToolConsultas1.AbrirDevConsulta 5, CnnPrincipal
        TxtCampos(5).Text = Principal.ToolConsultas1.DatSt
      
      Case 6
        Principal.ToolConsultas1.AbrirDevConsulta 3, CnnPrincipal
        TxtCampos(6).Text = Principal.ToolConsultas1.DatSt
      
    End Select
  End If
End Sub

Private Sub TxtCampos_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  Select Case Index
    Case 32
      If KeyAscii = 13 Then
        KeyAscii = 0
      End If
    Case 6, 7, 8, 9, 19, 20, 21, 22, 23, 24
      ValidarEntrada TxtCampos(Index), KeyAscii, 1
  End Select
  
End Sub
Private Sub TxtCampos_LostFocus(Index As Integer)
  TxtCampos(Index).BackColor = &H80000005
  TxtCampos(Index).ForeColor = &H80000012
End Sub

Private Sub TxtCampos_Validate(Index As Integer, Cancel As Boolean)
  Select Case Index
    Case 8, 9
      If Val(TxtCampos(Index).Text) <> 0 Then
        AbrirRecorset rstUniversal, "SELECT IdCiudad, NmCiudad, CodMinTrans  From Ciudades where IdCiudad=" & TxtCampos(Index), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          TxtNmCiudad(Index).Text = rstUniversal!NmCiudad & ""
          TxtCampos(Index).Tag = rstUniversal!CodMinTrans & ""
        Else
          TxtNmCiudad(Index).Text = "": TxtCampos(Index).Text = ""
        End If
        CerrarRecorset rstUniversal
      End If
    Case 7
      If Val(TxtCampos(Index).Text) <> 0 Then
        AbrirRecorset rstUniversal, "SELECT IdRuta, NmRuta From Rutas where IdRuta=" & TxtCampos(Index), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          TxtNmRuta.Text = rstUniversal!NmRuta & ""
        Else
          TxtNmRuta.Text = "": TxtCampos(Index).Text = ""
        End If
        CerrarRecorset rstUniversal
      End If
    Case 6
      If TxtCampos(6).Text <> "" Then
        AbrirRecorset rstUniversal, "Select IdConductor, Concat(Nombre, ' ', Apellido1,  ' ', Apellido2) as NmConductor, FhVenceLic, ConductorInactivo From Conductores where IdConductor='" & TxtCampos(6).Text & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          If CDate(rstUniversal.Fields("FhVenceLic")) < Date Then
            MsgTit "El conductor no puede viajar porque tiene la licencia vencida o esta inactivo"
            TxtCampos(6).Text = "": LblNmConductor.Caption = ""
          Else
            If Val(rstUniversal.Fields("ConductorInactivo")) = 0 Then
              LblNmConductor.Caption = rstUniversal.Fields("NmConductor")
            Else
              MsgTit "El conductor no puede viajar porque esta inactivo"
              TxtCampos(6).Text = "": LblNmConductor.Caption = ""
            End If
          End If
        Else
          MsgBox "El conductor no existe", vbCritical
          LblNmConductor.Caption = "": TxtCampos(6).Text = ""
        End If
        CerrarRecorset rstUniversal
      End If
    
    Case 5
      AbrirRecorset rstUniversal, "Select IdPlaca, VenceSoat, Inactivo from Vehiculos where IdPlaca='" & TxtCampos(5).Text & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.EOF = True Then
        MsgTit "El vehiculo no existe"
        TxtCampos(5).Text = ""
      Else
        If Val(rstUniversal.Fields("Inactivo")) = 1 Or CDate(rstUniversal.Fields("VenceSoat")) < Date Then
          MsgTit "El vehiculo esta inactivo o el soat esta vencido, no se puede despachar este vehiculo"
          TxtCampos(5).Text = ""
        End If
      End If
      CerrarRecorset rstUniversal
    
  End Select
  If Index >= 19 And Index <= 24 Then Calcular
End Sub
Private Sub Calcular()
  
  'TxtCampos(21).Text = Format(DctoPapeleria * Val(Format(TxtCampos(19).Text, "0;(0)")) / 100, "#,##0.00;(#,##0.00)")
  'TxtCampos(24).Text = Format(Estampilla, "#,##0.00;(#,##0.00)")
  
  'Retencion en la fuente
  If Val(Format(TxtCampos(19).Text, "0;(0)")) >= VrMayorA Then
        TxtCampos(27).Text = Format(RteFte * Val(Format(TxtCampos(19).Text, "0;(0)")) / 100, "#,##0.00;(#,##0.00)")
  Else
    TxtCampos(27).Text = 0
  End If
  'Industria y comercio
  TxtCampos(26).Text = Format(IndCom * Val(Format(TxtCampos(19).Text, "0;(0)")) / 100, "#,##0.00;(#,##0.00)")
  
  TxtCampos(28).Text = Format(Val(Format(TxtCampos(21).Text, "0;(0)")) + Val(Format(TxtCampos(22).Text, "0;(0)")) + Val(Format(TxtCampos(23).Text, "0;(0)")) + Val(Format(TxtCampos(24).Text, "0;(0)")), "#,##0.00;(#,##0.00)")
  TxtCampos(30).Text = Format(Val(Format(TxtCampos(19).Text, "0;(0)")) - (Val(Format(TxtCampos(26).Text, "0;(0)")) + Val(Format(TxtCampos(27).Text, "0;(0)")) + Val(Format(TxtCampos(20).Text, "0;(0)"))), "#,##0.00;(#,##0.00)")
  TxtCampos(29).Text = Format((Val(Format(TxtCampos(19).Text, "0;(0)")) + Val(Format(TxtCampos(25).Text, "0;(0)"))) - (Val(Format(TxtCampos(28).Text, "0;(0)")) + Val(Format(TxtCampos(20).Text, "0;(0)")) + Val(Format(TxtCampos(27).Text, "0;(0)")) + Val(Format(TxtCampos(26).Text, "0;(0)"))), "#,##0.00;(#,##0.00)")
End Sub

Private Function ValidarManElectronico() As Boolean
  Dim rstConductor As New ADODB.Recordset
  Dim rstVehiculo As New ADODB.Recordset
  rstConductor.CursorLocation = adUseClient
  rstVehiculo.CursorLocation = adUseClient
  
  ValidarManElectronico = False
  AbrirRecorset rstUniversal, "Select*from despachos where OrdDespacho=" & TxtCampos(1), CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.Fields("IdVehiculo") <> "" Then
    AbrirRecorset rstVehiculo, "Select*from vehiculos where IdPlaca='" & rstUniversal.Fields("IdVehiculo") & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstUniversal.Fields("IdConductor") <> "" Then
      AbrirRecorset rstConductor, "Select*from conductores where IdConductor='" & rstUniversal.Fields("IdConductor") & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstConductor.Fields("TpIdConductor") = "C" Or rstConductor.Fields("TpIdConductor") = "E" Or rstConductor.Fields("TpIdConductor") = "T" Then
        If rstVehiculo.Fields("VehConfiguracion") <> "" Then
          
          ValidarManElectronico = True
        Else
          MsgBox "El vehiculo debe tener una configuracion", vbCritical
        End If
      Else
        MsgBox "El documento del conductor debe ser cedula, estranjeria o tarjeta", vbCritical
      End If
    Else
      MsgBox "El despacho debe tener un conductor", vbCritical
    End If
  Else
    MsgBox "El despacho debe tener una placa", vbCritical
  End If
  CerrarRecorset rstUniversal
End Function

Private Sub TxtPagoCargue_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtPagoDescargue_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Cobros(intNumeroDespacho As Double)
  Dim strSql As String
  Dim rstPagos As New ADODB.Recordset
  Dim rstDespacho As New ADODB.Recordset
  Dim rstGuias As New ADODB.Recordset
  Dim douTotalCobroFleteDestino As Double
  Dim douTotalCobroManejoDestino As Double
  Dim douAbonos As Double
  rstGuias.CursorLocation = adUseClient
  rstDespacho.CursorLocation = adUseClient
  rstPagos.CursorLocation = adUseClient
  strSql = "SELECT Guia, VrFlete, VrManejo, Abonos FROM guias WHERE IdDespacho = " & intNumeroDespacho & " AND TipoCobro = 2"
  AbrirRecorset rstGuias, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
  douTotalCobroFleteDestino = 0
  douTotalCobroManejoDestino = 0
  douAbonos = 0
  Do While rstGuias.EOF = False
      douTotalCobroFleteDestino = douTotalCobroFleteDestino + Val(rstGuias!VrFlete)
      douTotalCobroManejoDestino = douTotalCobroManejoDestino + Val(rstGuias!VrManejo)
      If Val(rstGuias!Abonos) > 0 Then
        strSql = "SELECT VrFlete, VrManejo FROM recibos_caja_soporte WHERE Guia = " & rstGuias!Guia
        AbrirRecorset rstPagos, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
        Do While rstPagos.EOF = False
          douTotalCobroManejoDestino = douTotalCobroManejoDestino - Val(rstPagos!VrManejo)
          douTotalCobroFleteDestino = douTotalCobroFleteDestino - Val(rstPagos!VrFlete)
          rstPagos.MoveNext
        Loop
        CerrarRecorset rstPagos
        douAbonos = douAbonos + Val(rstGuias!Abonos)
      End If
    rstGuias.MoveNext
  Loop
  AbrirRecorset rstDespacho, "UPDATE Despachos SET ManejoCE = " & douTotalCobroManejoDestino & ", FleteCE = " & douTotalCobroFleteDestino & ", AbonosCE=" & douAbonos & ", FleteContado=0, ManejoContado=0, FleteCorriente=0, ManejoCorriente=0, FleteCETotal=0, ManejoCETotal=0 WHERE OrdDespacho = " & intNumeroDespacho, CnnPrincipal, adOpenDynamic, adLockOptimistic
  'Nueva actualizacion
  CerrarRecorset rstGuias
  strSql = "SELECT TipoCobro, SUM(VrFlete) as VrFlete, SUM(VrManejo) as VrManejo FROM guias WHERE IdDespacho = " & intNumeroDespacho & " GROUP BY TipoCobro"
  AbrirRecorset rstGuias, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
  Do While rstGuias.EOF = False
    If Val(rstGuias!TipoCobro) = 1 Then
      AbrirRecorset rstDespacho, "UPDATE Despachos SET FleteContado=" & rstGuias!VrFlete & ", ManejoContado=" & rstGuias!VrManejo & " WHERE OrdDespacho = " & intNumeroDespacho, CnnPrincipal, adOpenDynamic, adLockOptimistic
    End If
    If Val(rstGuias!TipoCobro) = 2 Then
      AbrirRecorset rstDespacho, "UPDATE Despachos SET FleteCETotal=" & rstGuias!VrFlete & ", ManejoCETotal=" & rstGuias!VrManejo & " WHERE OrdDespacho = " & intNumeroDespacho, CnnPrincipal, adOpenDynamic, adLockOptimistic
    End If
    If Val(rstGuias!TipoCobro) = 3 Then
      AbrirRecorset rstDespacho, "UPDATE Despachos SET FleteCorriente=" & rstGuias!VrFlete & ", ManejoCorriente=" & rstGuias!VrManejo & " WHERE OrdDespacho = " & intNumeroDespacho, CnnPrincipal, adOpenDynamic, adLockOptimistic
    End If
    rstGuias.MoveNext
  Loop
  CerrarRecorset rstGuias
End Sub
