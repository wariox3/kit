VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmCerrarDespacho 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cerrar manifiesto..."
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmContraentregas 
      Caption         =   "Contraentregas"
      Height          =   2055
      Left            =   8520
      TabIndex        =   46
      Top             =   3960
      Width           =   3015
      Begin MSMask.MaskEdBox TxtValores 
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
         Index           =   14
         Left            =   1560
         TabIndex        =   47
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtValores 
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
         Index           =   15
         Left            =   1560
         TabIndex        =   49
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtValores 
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
         Index           =   16
         Left            =   1560
         TabIndex        =   51
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtValores 
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
         Index           =   17
         Left            =   1560
         TabIndex        =   53
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtValores 
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
         Index           =   18
         Left            =   1560
         TabIndex        =   57
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "No pagados:"
         Height          =   195
         Index           =   4
         Left            =   465
         TabIndex        =   58
         Top             =   600
         Width           =   915
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Recaudos:"
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   54
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Flete:"
         Height          =   195
         Index           =   2
         Left            =   990
         TabIndex        =   52
         Top             =   975
         Width           =   390
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Manejo:"
         Height          =   195
         Index           =   1
         Left            =   810
         TabIndex        =   50
         Top             =   1335
         Width           =   570
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Abonos:"
         Height          =   195
         Index           =   0
         Left            =   795
         TabIndex        =   48
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame FraConceptos 
      Height          =   1575
      Left            =   120
      TabIndex        =   38
      Top             =   2280
      Width           =   11415
      Begin VB.TextBox TxtNmConcepto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   42
         Top             =   240
         Width           =   6495
      End
      Begin VB.TextBox TxtIdConcepto 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TxtValor 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TxtComentario 
         Height          =   525
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   10215
      End
      Begin VB.CommandButton CmdQuitar 
         Caption         =   "Quitar"
         Height          =   255
         Left            =   2760
         TabIndex        =   39
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Concepto:"
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Left            =   600
         TabIndex        =   41
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Comentarios:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   915
      End
   End
   Begin VB.CommandButton CmdCerrarDespacho 
      Caption         =   "Cerrar despacho"
      Height          =   375
      Left            =   7200
      TabIndex        =   37
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton CmdAyuda 
      Caption         =   "Ayuda"
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Frame FraFlete 
      Caption         =   "Flete y anticipo"
      Enabled         =   0   'False
      Height          =   975
      Left            =   120
      TabIndex        =   31
      Top             =   3960
      Width           =   2535
      Begin MSMask.MaskEdBox TxtValores 
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
         Index           =   0
         Left            =   1200
         TabIndex        =   32
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtValores 
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
         Index           =   1
         Left            =   1200
         TabIndex        =   33
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Anticipo:"
         Height          =   195
         Index           =   36
         Left            =   480
         TabIndex        =   35
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "Valor Flete:"
         Height          =   195
         Left            =   360
         TabIndex        =   34
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame FraDescuentos 
      Caption         =   "Otros descuentos"
      Height          =   2055
      Left            =   2760
      TabIndex        =   24
      Top             =   3960
      Width           =   2415
      Begin MSMask.MaskEdBox TxtValores 
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
         Index           =   7
         Left            =   960
         TabIndex        =   7
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtValores 
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
         Index           =   6
         Left            =   960
         TabIndex        =   6
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtValores 
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
         Index           =   5
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtValores 
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
         Index           =   4
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtValores 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   960
         TabIndex        =   25
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Papeleria:"
         Height          =   195
         Index           =   32
         Left            =   210
         TabIndex        =   30
         Top             =   240
         Width           =   705
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Seguridad:"
         Height          =   195
         Index           =   33
         Left            =   150
         TabIndex        =   29
         Top             =   600
         Width           =   765
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Estampilla:"
         Height          =   195
         Index           =   35
         Left            =   165
         TabIndex        =   28
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Cargue:"
         Height          =   195
         Index           =   34
         Left            =   360
         TabIndex        =   27
         Top             =   960
         Width           =   555
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Index           =   40
         Left            =   480
         TabIndex        =   26
         Top             =   1680
         Width           =   405
      End
   End
   Begin VB.Frame FraTotales 
      Caption         =   "Resumen"
      Enabled         =   0   'False
      Height          =   2055
      Left            =   5280
      TabIndex        =   15
      Top             =   3960
      Width           =   3135
      Begin MSMask.MaskEdBox TxtValores 
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
         Index           =   9
         Left            =   1680
         TabIndex        =   16
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtValores 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   1680
         TabIndex        =   17
         Top             =   1335
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtValores 
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
         Index           =   13
         Left            =   1680
         TabIndex        =   18
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtValores 
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
         Index           =   11
         Left            =   1680
         TabIndex        =   19
         Top             =   945
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtValores 
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
         Index           =   10
         Left            =   1680
         TabIndex        =   55
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Contra Entrega:"
         Height          =   195
         Index           =   39
         Left            =   480
         TabIndex        =   56
         Top             =   615
         Width           =   1110
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Flete adicional:"
         Height          =   195
         Index           =   41
         Left            =   555
         TabIndex        =   23
         Top             =   967
         Width           =   1065
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Descuentos Varios:"
         Height          =   195
         Index           =   31
         Left            =   240
         TabIndex        =   22
         Top             =   1680
         Width           =   1380
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Manifiesto:"
         Height          =   195
         Index           =   46
         Left            =   405
         TabIndex        =   21
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Total Viaje:"
         Height          =   195
         Index           =   23
         Left            =   825
         TabIndex        =   20
         Top             =   1323
         Width           =   795
      End
   End
   Begin VB.Frame FraIndCom 
      Caption         =   "Rte Fuente/Ind Comercio"
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   5040
      Width           =   2535
      Begin MSMask.MaskEdBox TxtValores 
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
         Index           =   3
         Left            =   1185
         TabIndex        =   11
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtValores 
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
         Index           =   2
         Left            =   1185
         TabIndex        =   12
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Ret. Fuente:"
         Height          =   195
         Index           =   29
         Left            =   255
         TabIndex        =   14
         Top             =   240
         Width           =   885
      End
      Begin VB.Label LblUniversal 
         AutoSize        =   -1  'True
         Caption         =   "Ind. Comercio:"
         Height          =   195
         Index           =   30
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1020
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   9360
      TabIndex        =   9
      Top             =   6240
      Width           =   2175
   End
   Begin MSComctlLib.ListView LstConceptos 
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Concepto"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tp"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Valor"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Comentarios"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.Label LblIdDespacho 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   195
      Left            =   1080
      TabIndex        =   45
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label3 
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
      Left            =   120
      TabIndex        =   44
      Top             =   120
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   120
      X2              =   11520
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      Index           =   1
      X1              =   11520
      X2              =   120
      Y1              =   6120
      Y2              =   6120
   End
End
Attribute VB_Name = "FrmCerrarDespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstInsertar As New ADODB.Recordset


Private Sub CmdAceptar_Click()
  If Val(Me.Tag) = 0 Then
    II = MsgBox("¿Desea guardar los cambios que le hizo al manifiesto de carga?", vbQuestion + vbYesNoCancel, "¿Desea guardar los cambios?")
    Select Case II
      Case 6 ' Si
        GuardarCambios
        Unload Me
      Case 7 ' No
        Unload Me
      Case 2 ' Cancelar
    End Select
  Else
    Unload Me
  End If
End Sub
Private Sub GuardarCambios()
  AbrirRecorset rstInsertar, "Update despachos set" & _
  " SaldoDesp=" & Val(TxtValores(9)) & ", VrFleteAdicional=" & Val(TxtValores(11)) & "," & _
  " TotalViaje=" & Val(TxtValores(12)) & ", VrOtrosDctos=" & Val(TxtValores(13)) & "," & _
  " VrDctoPapeleria=" & Val(TxtValores(4)) & ", VrDctoSeguridad=" & Val(TxtValores(5)) & "," & _
  " VrDctoCargue=" & Val(TxtValores(6)) & ", VrDctoEstampilla=" & Val(TxtValores(7)) & "," & _
  " AbonosCE=" & Val(TxtValores(14)) & ", FletesNoCancelados=" & Val(TxtValores(17)) & _
  " where OrdDespacho=" & Val(LblIdDespacho.Caption), CnnPrincipal, adOpenDynamic, adLockOptimistic
End Sub
Private Sub CmdAgregar_Click()
  AbrirRecorset rstUniversal, "Select IdConcepto, NmConcepto, Tipo from conceptoscontables where IdConcepto=" & Val(TxtIdConcepto.Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.RecordCount > 0 Then
    If Val(TxtValor.Text) <> 0 Then
        AbrirRecorset rstInsertar, "insert into despachosconceptos (IdDespacho, IdConcepto, Tipo, Valor, Comentarios) values(" & Val(LblIdDespacho.Caption) & ", " & Val(TxtIdConcepto.Text) & ", " & Val(rstUniversal.Fields("Tipo")) & ", " & Val(TxtValor.Text) & ", '" & TxtComentario.Text & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
        
      If Val(rstUniversal.Fields("Tipo")) = 2 Then
        TxtValores(11) = Val(TxtValores(11)) + Val(TxtValor.Text)
      Else
        TxtValores(13) = Val(TxtValores(13)) + Val(TxtValor.Text)
      End If
      TxtIdConcepto.Text = ""
      TxtNmConcepto.Text = ""
      TxtValor.Text = ""
      TxtComentario.Text = ""
      Calcular
      GuardarCambios
      VerConceptos
      TxtIdConcepto.SetFocus
    Else
      MsgBox "No puede ingresar un valor en ceros", vbCritical, "Valor no permitido": TxtValor.SetFocus
    End If
  End If
  CerrarRecorset rstUniversal

End Sub

Private Sub CmdCerrarDespacho_Click()
  If MsgBox("¿Esta seguro que desea cerrar la liquidacion de este despacho?", vbQuestion + vbYesNo, "¿Esta seguro de cerrar el viaje?") = vbYes Then
    AbrirRecorset rstUniversal, "Update Despachos set Cerrado=1 Where OrdDespacho=" & Val(LblIdDespacho.Caption), CnnPrincipal, adOpenDynamic, adLockOptimistic
    GuardarCambios
    Unload Me
  End If
End Sub

Private Sub CmdQuitar_Click()
  II = 1
  While II <= LstConceptos.ListItems.Count
    If LstConceptos.ListItems(II).Checked = True Then
    
      If LstConceptos.ListItems(II).SubItems(3) = 1 Then
        TxtValores(13) = Val(TxtValores(13)) - Val(LstConceptos.ListItems(II).SubItems(4))
      Else
        TxtValores(11) = Val(TxtValores(11)) - Val(LstConceptos.ListItems(II).SubItems(4))
      End If
      AbrirRecorset rstInsertar, "delete from despachosconceptos where ID=" & LstConceptos.ListItems(II).Text, CnnPrincipal, adOpenDynamic, adLockOptimistic
      LstConceptos.ListItems.Remove (II)
    Else
     II = II + 1
    End If
  Wend
  Calcular
  GuardarCambios
End Sub

Private Sub Form_Load()
  LblIdDespacho.Caption = FufuLo
  Me.Tag = II
  If Val(Me.Tag) = 1 Then
    FraConceptos.Enabled = False
    FraConceptos.Visible = False
    LstConceptos.Height = 3015
    CmdCerrarDespacho.Enabled = False
  End If
  AbrirRecorset rstUniversal, "Select*from slq_im_formatoliqdesp where OrdDespacho=" & Val(LblIdDespacho.Caption), CnnPrincipal, adOpenKeyset, adLockReadOnly
    If rstUniversal.EOF = False Then
      TxtValores(0) = rstUniversal!VrFlete
      TxtValores(1) = rstUniversal!VrAnticipo
      TxtValores(2) = rstUniversal!VrDctoRteFte
      TxtValores(3) = rstUniversal!VrDctoIndCom
      TxtValores(4) = rstUniversal!VrDctoPapeleria
      TxtValores(5) = rstUniversal!VrDctoSeguridad
      TxtValores(6) = rstUniversal!VrDctoCargue
      TxtValores(7) = rstUniversal!VrDctoEstampilla
      TxtValores(16) = rstUniversal!FleteCE
      TxtValores(15) = rstUniversal!ManejoCE
      TxtValores(18) = rstUniversal!TRecaudo
      TxtValores(14) = rstUniversal!AbonosCE
      TxtValores(17) = rstUniversal!FletesNoCancelados
      TxtValores(10) = rstUniversal!FleteCE + rstUniversal!ManejoCE + rstUniversal!TRecaudo
      TxtValores(11) = rstUniversal!VrFleteAdicional
      TxtValores(13) = rstUniversal!VrOtrosDctos
      Calcular
    End If
  CerrarRecorset rstUniversal
  AbrirRecorset rstUniversal, "Select*from DespachosConceptos where IdDespacho=" & Val(LblIdDespacho.Caption), CnnPrincipal, adOpenKeyset, adLockReadOnly
    Do While rstUniversal.EOF = False
      Set Item = LstConceptos.ListItems.Add(, , rstUniversal!Id)
      Item.SubItems(1) = rstUniversal!IdConcepto
      If rstUniversal!IdConcepto = 0 Then Item.SubItems(2) = "FLETE ADICIONAL" Else Item.SubItems(2) = "DESCUENTO"
      Item.SubItems(3) = rstUniversal!Valor
      Item.SubItems(4) = rstUniversal!Comentarios
      rstUniversal.MoveNext
    Loop
  CerrarRecorset rstUniversal
  VerConceptos
End Sub
Sub Calcular()
  TxtValores(8) = Val(TxtValores(4)) + Val(TxtValores(5)) + Val(TxtValores(6)) + Val(TxtValores(7))
  TxtValores(9) = Val(TxtValores(0)) - (Val(TxtValores(1)) + Val(TxtValores(2)) + Val(TxtValores(3))) - (Val(TxtValores(8)))
  TxtValores(12) = ((Val(TxtValores(0)) + Val(TxtValores(11))) - (Val(TxtValores(1)) + Val(TxtValores(2)) + Val(TxtValores(3)) + Val(TxtValores(8)) + Val(TxtValores(10)) + Val(TxtValores(13))) - Val(TxtValores(14)) + Val(TxtValores(17)))
End Sub
Sub VerConceptos()
  LstConceptos.ListItems.Clear
  AbrirRecorset rstUniversal, "Select despachosconceptos.*, NmConcepto from despachosconceptos left join conceptoscontables on despachosconceptos.IdConcepto=conceptoscontables.IdConcepto where IdDespacho=" & Val(LblIdDespacho.Caption), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      Do While rstUniversal.EOF = False
        Set Item = LstConceptos.ListItems.Add(, , rstUniversal!Id)
          Item.SubItems(1) = rstUniversal!IdConcepto
          Item.SubItems(2) = rstUniversal!NmConcepto
          Item.SubItems(3) = rstUniversal!Tipo
          Item.SubItems(4) = rstUniversal!Valor
          Item.SubItems(5) = rstUniversal!Comentarios & ""
        rstUniversal.MoveNext
      Loop
    End If
  CerrarRecorset rstUniversal
End Sub
Private Sub TxtComentario_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys vbTab
  End If
End Sub



Private Sub TxtIdConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    FrmBuscarConcepto.Show 1
    If FufuLo <> 0 Then
      AbrirRecorset rstUniversal, "Select IdConcepto, NmConcepto from conceptoscontables where IdConcepto=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.RecordCount > 0 Then
        TxtNmConcepto.Text = rstUniversal.Fields("NmConcepto")
        TxtIdConcepto.Text = FufuLo
      End If
      CerrarRecorset rstUniversal
    Else
      TxtNmConcepto.Text = ""
      TxtIdConcepto.Text = ""
    End If
  End If
End Sub

Private Sub TxtIdConcepto_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtValor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtValores_KeyPress(Index As Integer, KeyAscii As Integer)

  Select Case Index
    Case 4, 5, 6, 7, 14, 17
      ValidarEntrada TxtValores(Index), KeyAscii, 1
  End Select
  If KeyAscii = 13 Then
      SendKeys vbTab
  End If
End Sub

Private Sub TxtValores_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  Calcular
End Sub
