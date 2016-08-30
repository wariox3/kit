VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmFacturas 
   Caption         =   "Facturas..."
   ClientHeight    =   11070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14880
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11070
   ScaleWidth      =   14880
   WindowState     =   2  'Maximized
   Begin VB.Frame FraRecuento 
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      TabIndex        =   62
      Top             =   8160
      Width           =   10335
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   19
         Left            =   7680
         TabIndex        =   68
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   18
         Left            =   4800
         TabIndex        =   66
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   17
         Left            =   1800
         TabIndex        =   64
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nro Conceptos:"
         Height          =   195
         Left            =   6480
         TabIndex        =   67
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nro Planillas:"
         Height          =   195
         Left            =   3840
         TabIndex        =   65
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nro Guias:"
         Height          =   195
         Left            =   960
         TabIndex        =   63
         Top             =   240
         Width           =   750
      End
   End
   Begin TabDlg.SSTab SSTDatos 
      Height          =   5295
      Left            =   240
      TabIndex        =   5
      Tag             =   "Vacia"
      Top             =   1320
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Datos Basicos"
      TabPicture(0)   =   "FrmFacturas.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GrillaFacturadas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraCliente"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraNotas"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdVerGuias"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmdAgregarQuitar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CmdLiberarGuias"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Planillas"
      TabPicture(1)   =   "FrmFacturas.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraPlanillas"
      Tab(1).Control(1)=   "CmdVerPlanillas"
      Tab(1).Control(2)=   "CmdManPlanillas"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "&Otros Conceptos"
      TabPicture(2)   =   "FrmFacturas.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FraConceptos"
      Tab(2).Control(1)=   "CmdVerConceptos"
      Tab(2).Control(2)=   "CmdMantenimientoConceptos"
      Tab(2).ControlCount=   3
      Begin VB.CommandButton CmdLiberarGuias 
         Caption         =   "Liberar guias"
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   4920
         Width           =   2055
      End
      Begin VB.CommandButton CmdMantenimientoConceptos 
         Caption         =   "Agergar / Quitar Conceptos"
         Height          =   255
         Left            =   -74760
         TabIndex        =   74
         Top             =   4320
         Width           =   2535
      End
      Begin VB.CommandButton CmdVerConceptos 
         Caption         =   "Ver conceptos"
         Height          =   255
         Left            =   -68520
         TabIndex        =   73
         Top             =   4320
         Width           =   2535
      End
      Begin VB.CommandButton CmdManPlanillas 
         Caption         =   "Agregar / Quitar planillas"
         Height          =   255
         Left            =   -74640
         TabIndex        =   72
         Top             =   4200
         Width           =   2295
      End
      Begin VB.CommandButton CmdVerPlanillas 
         Caption         =   "Ver Planillas"
         Height          =   255
         Left            =   -67680
         TabIndex        =   71
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Frame FraPlanillas 
         Caption         =   "Planillas"
         Enabled         =   0   'False
         Height          =   3675
         Left            =   -74760
         TabIndex        =   56
         Top             =   480
         Width           =   8775
         Begin VB.CommandButton CmdCambiarNPlanilla 
            Caption         =   "Cambiar"
            Height          =   255
            Left            =   4080
            TabIndex        =   76
            Top             =   3720
            Width           =   975
         End
         Begin VB.CommandButton CmdQuitarPlanilla 
            Caption         =   "Quitar"
            Height          =   255
            Left            =   3120
            TabIndex        =   70
            Top             =   3720
            Width           =   855
         End
         Begin VB.CommandButton CmdCargarRemisionesPlanilla 
            Caption         =   "Cargar guias"
            Height          =   255
            Left            =   5400
            TabIndex        =   61
            Top             =   3720
            Width           =   1575
         End
         Begin VB.CommandButton CmdAgregarPlanilla 
            Caption         =   "Agregar"
            Height          =   255
            Left            =   2280
            TabIndex        =   60
            Top             =   3720
            Width           =   855
         End
         Begin VB.TextBox TxtRelCliente 
            Height          =   285
            Left            =   1080
            TabIndex        =   59
            Top             =   3720
            Width           =   1095
         End
         Begin MSComctlLib.ListView LstPlanillas 
            Height          =   3255
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   5741
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Planilla"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Rel Cliente"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Flete"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Manejo"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "NroRemisiones"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Rel Cliente:"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   3720
            Width           =   810
         End
      End
      Begin VB.Frame FraConceptos 
         Caption         =   "Conceptos"
         Enabled         =   0   'False
         Height          =   3000
         Left            =   -74760
         TabIndex        =   49
         Top             =   480
         Width           =   8775
         Begin VB.TextBox TxtIdConcepto 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6360
            TabIndex        =   75
            Top             =   3360
            Width           =   495
         End
         Begin VB.TextBox TxtValorConcepto 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7560
            TabIndex        =   52
            Top             =   3000
            Width           =   1095
         End
         Begin VB.CommandButton CmdAgregar 
            Caption         =   "Agregar"
            Height          =   255
            Left            =   960
            TabIndex        =   53
            Top             =   3360
            Width           =   1935
         End
         Begin VB.CommandButton CmdQuitar 
            Caption         =   "Quitar"
            Height          =   255
            Left            =   3000
            TabIndex        =   57
            Top             =   3360
            Width           =   1935
         End
         Begin MSComctlLib.ListView LstConceptos 
            Height          =   2655
            Left            =   120
            TabIndex        =   50
            Tag             =   "Vacia"
            Top             =   240
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   4683
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Id"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Concepto"
               Object.Width           =   10583
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Valor"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSDataListLib.DataCombo CboConceptos 
            Height          =   315
            Left            =   960
            TabIndex        =   51
            Top             =   3000
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Height          =   195
            Left            =   7080
            TabIndex        =   55
            Top             =   3000
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Concepto:"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   3000
            Width           =   735
         End
      End
      Begin VB.CommandButton CmdAgregarQuitar 
         Caption         =   "Agregar/Quitar guias >>"
         Height          =   255
         Left            =   8160
         TabIndex        =   37
         Top             =   4920
         Width           =   2055
      End
      Begin VB.CommandButton CmdVerGuias 
         Caption         =   "Ver Guias"
         Height          =   255
         Left            =   6000
         TabIndex        =   36
         Top             =   4920
         Width           =   2055
      End
      Begin VB.Frame FraNotas 
         Caption         =   "Notas adicionales"
         Enabled         =   0   'False
         Height          =   855
         Left            =   240
         TabIndex        =   33
         Top             =   1680
         Width           =   9975
         Begin VB.TextBox TxtCampos 
            Height          =   495
            Index           =   5
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   240
            Width           =   9735
         End
      End
      Begin VB.Frame FraCliente 
         Enabled         =   0   'False
         Height          =   1335
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   9975
         Begin VB.TextBox TxtCampos 
            Height          =   285
            Index           =   22
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox TxtCampos 
            Height          =   285
            Index           =   21
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox TxtCampos 
            Height          =   285
            Index           =   20
            Left            =   8640
            TabIndex        =   2
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox TxtNmCliente 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3360
            TabIndex        =   7
            Top             =   240
            Width           =   5895
         End
         Begin VB.TextBox TxtCampos 
            Height          =   285
            Index           =   3
            Left            =   1800
            TabIndex        =   0
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Centro operaciones:"
            Height          =   195
            Left            =   360
            TabIndex        =   82
            Top             =   960
            Width           =   1425
         End
         Begin VB.Label LblNmCentroOperaciones 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3360
            TabIndex        =   81
            Top             =   960
            Width           =   4575
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Plazo:"
            Height          =   195
            Left            =   8040
            TabIndex        =   79
            Top             =   600
            Width           =   435
         End
         Begin VB.Label LblFormaPago 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3360
            TabIndex        =   78
            Top             =   600
            Width           =   4575
         End
         Begin VB.Label LblIdFormaPago 
            AutoSize        =   -1  'True
            Caption         =   "Forma pago:"
            Height          =   195
            Left            =   840
            TabIndex        =   77
            Top             =   600
            Width           =   885
         End
         Begin VB.Label LblTitulo 
            AutoSize        =   -1  'True
            Caption         =   "ID Cliente:"
            Height          =   195
            Index           =   0
            Left            =   960
            TabIndex        =   8
            Top             =   240
            Width           =   735
         End
      End
      Begin MSDataGridLib.DataGrid GrillaFacturadas 
         Height          =   2175
         Left            =   240
         TabIndex        =   38
         Top             =   2640
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   3836
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
         ColumnCount     =   9
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
            DataField       =   "FhEntradaBodega"
            Caption         =   "FhEntra"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/mm/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   9226
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "DocCliente"
            Caption         =   "Doc Cliente"
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
            DataField       =   "NmCiudad"
            Caption         =   "Destino"
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
            DataField       =   "Unidades"
            Caption         =   "Und"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "KilosFacturados"
            Caption         =   "K Fact"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "VrFlete"
            Caption         =   "Flete"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00;(#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "VrManejo"
            Caption         =   "Manejo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00;(#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "VrDeclarado"
            Caption         =   "Declarado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00;(#,##0.00)"
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
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1154.835
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraEstCuenta 
      Caption         =   "Estado de cuenta"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   8160
      TabIndex        =   24
      Top             =   6720
      Width           =   2415
      Begin VB.TextBox TxtCampos 
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
         Index           =   15
         Left            =   840
         TabIndex        =   47
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   16
         Left            =   840
         TabIndex        =   31
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtCampos 
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
         Left            =   840
         TabIndex        =   30
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   46
         Top             =   960
         Width           =   405
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Width           =   450
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Abono:"
         Height          =   195
         Index           =   12
         Left            =   180
         TabIndex        =   29
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Frame FraDescuentos 
      Caption         =   "Descuentos / Total"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   3360
      TabIndex        =   23
      Top             =   6720
      Width           =   3975
      Begin VB.TextBox TxtCampos 
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
         Left            =   3240
         TabIndex        =   45
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   10
         Left            =   3240
         TabIndex        =   43
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TxtCampos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   3240
         TabIndex        =   34
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   11
         Left            =   960
         TabIndex        =   28
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   9
         Left            =   960
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Base:"
         Height          =   195
         Left            =   2760
         TabIndex        =   44
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Base:"
         Height          =   195
         Left            =   2790
         TabIndex        =   42
         Top             =   240
         Width           =   405
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Antes de:"
         Height          =   195
         Index           =   9
         Left            =   2520
         TabIndex        =   35
         Top             =   960
         Width           =   675
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Comercial:"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Financiero:"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   780
      End
   End
   Begin VB.Frame FraCobros 
      Caption         =   "Valor cobros"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   240
      TabIndex        =   14
      Top             =   6720
      Width           =   2415
      Begin VB.TextBox TxtCampos 
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
         Index           =   6
         Left            =   840
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   7
         Left            =   840
         TabIndex        =   16
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtCampos 
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
         Index           =   8
         Left            =   840
         TabIndex        =   15
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Flete:"
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   20
         Top             =   240
         Width           =   390
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Manejo:"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   19
         Top             =   600
         Width           =   570
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Otros:"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   18
         Top             =   960
         Width           =   420
      End
   End
   Begin VB.Frame FraEncabezado 
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   10335
      Begin VB.TextBox TxtEstado 
         Height          =   285
         Left            =   7680
         TabIndex        =   48
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox TxtCampos 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   7320
         TabIndex        =   40
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   22
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   2
         Left            =   5520
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   8
         Left            =   6720
         TabIndex        =   41
         Top             =   240
         Width           =   540
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Vence:"
         Height          =   195
         Index           =   7
         Left            =   4920
         TabIndex        =   13
         Top             =   240
         Width           =   510
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Index           =   6
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Factura:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   585
      End
   End
   Begin MSComctlLib.Toolbar ToolFacturas 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   1005
      ButtonWidth     =   900
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
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Imp1"
                  Text            =   "Imprimir por guias"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Imp2"
                  Text            =   "Imprimir por planillas"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Imp3"
                  Text            =   "Imprimir por conceptos"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Imp4"
                  Text            =   "Imprimir por guias formato"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Imp5"
                  Text            =   "Imprimir por planillas formato"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Imp6"
                  Text            =   "Imprimir por conceptos formato"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Car"
            Object.ToolTipText     =   "Carga informacion adicional [Pausa]"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Otros"
            Object.ToolTipText     =   "Otras utilidades"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Otr1"
                  Text            =   "Corregir guias"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Editando As Boolean
Dim rstFacturas As New ADODB.Recordset
Dim rstTem As New ADODB.Recordset
Dim strSqlFacturas As String

Private Sub CboConceptos_GotFocus()
  Dim rstConceptos As New ADODB.Recordset
  rstConceptos.CursorLocation = adUseClient
  AbrirRecorset rstConceptos, "Select IdConcepto, NmConcepto from ConceptosContables order by NmConcepto", CnnPrincipal, adOpenDynamic, adLockOptimistic
  CboConceptos.ListField = "NmConcepto"
  Set CboConceptos.RowSource = rstConceptos
End Sub

Private Sub CboConceptos_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then TxtValorConcepto.SetFocus
End Sub

Private Sub CboConceptos_LostFocus()
  AbrirRecorset rstUniversal, "Select IdConcepto, NmConcepto from ConceptosContables where NmCOncepto='" & CboConceptos & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = False Then
    TxtIdConcepto.Text = rstUniversal!IdConcepto
  Else
    TxtIdConcepto.Text = ""
    CboConceptos.Text = ""
  End If
  CerrarRecorset rstUniversal
End Sub


Private Sub CmdAgregar_Click()
  AbrirRecorset rstUniversal, "Select IdFactura, Estado, NroGuias, NroPlanillas from facturas where Estado='D' and NroPlanillas=0 and NroGuias=0 and IdFactura=" & Val(TxtCampos(0)), CnnPrincipal, adOpenDynamic, adLockOptimistic
    FufuLo = rstUniversal.RecordCount
  CerrarRecorset rstUniversal
  If FufuLo >= 1 Then
    If Val(TxtIdConcepto.Text) <> 0 Then
      If Val(TxtValorConcepto.Text) > 0 Then
        AbrirRecorset rstUniversal, "INSERT INTO ConceptosFacturas (IdFactura, IdConcepto, Valor) VALUES (" & Val(TxtCampos(0)) & ", " & Val(TxtIdConcepto.Text) & ", " & Val(TxtValorConcepto) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
        VerConceptos
        TxtCampos(8).Text = Val(Format(TxtCampos(8).Text, "0;(0)")) + Val(TxtValorConcepto)
        TxtCampos(19) = Val(TxtCampos(19)) + 1
        Liquidar
        AbrirRecorset rstUniversal, "Update Facturas set NroConceptos=NroConceptos+1, DctoComercial=" & Val(Format(TxtCampos(9).Text, "0;(0)")) & ", DctoFinanciero=" & Val(Format(TxtCampos(11).Text, "0;(0)")) & ", TotalFactura=" & Val(Format(TxtCampos(15).Text, "0;(0)")) & ", TOtros=" & Val(Format(TxtCampos(8).Text, "0;(0)")) & ", Saldo=" & Val(Format(TxtCampos(16).Text, "0;(0)")) & " where IdFactura=" & TxtCampos(0).Text, CnnPrincipal, adOpenDynamic, adLockOptimistic
        MsgBox "El concepto fue agregado con exito", vbExclamation
        TxtValorConcepto.Text = ""
        CboConceptos.Text = ""
        TxtIdConcepto.Text = ""
        AccionTool 17
        CboConceptos.SetFocus
      Else
        MsgBox "Debe digitar un valor para el concepto", vbCritical: TxtValorConcepto.SetFocus
      End If
    Else
      MsgBox "Debe digitar un concepto", vbCritical: CboConceptos.SetFocus
    End If
  Else
    MsgBox "La factura debe esta digitada, no puede tener remisiones asiganadas ni planillas", vbCritical
  End If
End Sub

Private Sub CmdAgregarPlanilla_Click()
  AbrirRecorset rstUniversal, "Select IdFactura, Estado, NroConceptos, NroGuias from facturas where Estado='D' and NroConceptos=0 and NroGuias=0 and IdFactura=" & Val(TxtCampos(0)), CnnPrincipal, adOpenDynamic, adLockOptimistic
    FufuLo = rstUniversal.RecordCount
  CerrarRecorset rstUniversal
  If FufuLo >= 1 Then
    If TxtRelCliente.Text <> "" Then
      AbrirRecorset rstUniversal, "INSERT INTO FacturasPlanillas (IdFactura, RelCliente, VrFletePlanilla, VrManejoPlanilla, NroGuiasPlanilla) VALUES (" & Val(TxtCampos(0)) & ", '" & TxtRelCliente.Text & "', 0, 0, 0)", CnnPrincipal, adOpenDynamic, adLockOptimistic
      VerPlanillas
      TxtRelCliente.Text = ""
      TxtCampos(18).Text = Val(TxtCampos(18)) + 1
      AbrirRecorset rstUniversal, "Update facturas set NroPlanillas=NroPlanillas+1 where IdFactura=" & Val(TxtCampos(0)), CnnPrincipal, adOpenDynamic, adLockOptimistic
      MsgBox "Planilla agregada con exito", vbExclamation
      AccionTool 17
    Else
      MsgBox "Debe especificar un numero de relacion de cliente", vbCritical
    End If
  Else
    MsgBox "La factura debe esta digitada, no puede tener remisiones asiganadas ni conceptos", vbCritical
  End If
End Sub

Private Sub CmdAgregarQuitar_Click()
  AbrirRecorset rstUniversal, "Select IdFactura, NroPlanillas, NroConceptos, Estado from facturas where NroPlanillas=0 and NroConceptos=0 and Estado='D' and IdFactura=" & Val(TxtCampos(0)), CnnPrincipal, adOpenDynamic, adLockOptimistic
  FufuLo = rstUniversal.RecordCount
  CerrarRecorset rstUniversal
  If FufuLo >= 1 Then
    AccionTool 19
    FrmLlenarFacturas.Show 1
    Liquidar
    AbrirRecorset rstUniversal, "Update Facturas set NroGuias=" & Val(TxtCampos(17).Text) & ", TFlete=" & Val(Format(TxtCampos(6).Text, "0;(0)")) & ", TManejo=" & Val(Format(TxtCampos(7).Text, "0;(0)")) & ", DctoComercial=" & Val(Format(TxtCampos(9).Text, "0;(0)")) & ", DctoFinanciero=" & Val(Format(TxtCampos(11).Text, "0;(0)")) & ", TotalFactura=" & Val(Format(TxtCampos(15).Text, "0;(0)")) & ", Saldo=" & Val(Format(TxtCampos(16).Text, "0;(0)")) & " where IdFactura=" & TxtCampos(0).Text, CnnPrincipal, adOpenDynamic, adLockOptimistic
    AccionTool 17
  Else
    MsgBox "No se le pueden agregar remisiones a la factura si esta ya esta impresa o anulada o si tiene planillas", vbCritical, "No se pueden agregar remisiones"
  End If
End Sub

Private Sub CmdCambiarNPlanilla_Click()
  Dim NRelacion As String
  NRelacion = InputBox("Digite el nuevo numero de planilla. Numero Anterior:" & LstPlanillas.ListItems(LstPlanillas.SelectedItem.Index).SubItems(1), "Nuevo numero de relacion")
  If Val(NRelacion) <> 0 Then
    AbrirRecorset rstUniversal, "Update facturasplanillas set RelCliente =" & NRelacion & " where IdPlanilla=" & Val(LstPlanillas.ListItems(LstPlanillas.SelectedItem.Index)), CnnPrincipal, adOpenDynamic, adLockOptimistic
  End If
  VerPlanillas
End Sub

Private Sub CmdCargarRemisionesPlanilla_Click()
  If LstPlanillas.ListItems.Count >= 1 Then
    If CpEstFactura(Val(TxtCampos(0))) = "D" Then
      AccionTool 19
      FufuLo = LstPlanillas.ListItems(LstPlanillas.SelectedItem.Index).Text
      FrmLlenarFacturasPlanilla.Show 1
      Liquidar
      VerPlanillas
      AbrirRecorset rstUniversal, "Update Facturas set TFlete=" & Val(Format(TxtCampos(6).Text, "0;(0)")) & ", TManejo=" & Val(Format(TxtCampos(7).Text, "0;(0)")) & ", DctoComercial=" & Val(Format(TxtCampos(9).Text, "0;(0)")) & ", DctoFinanciero=" & Val(Format(TxtCampos(11).Text, "0;(0)")) & ", TotalFactura=" & Val(Format(TxtCampos(15).Text, "0;(0)")) & ", Saldo=" & Val(Format(TxtCampos(16).Text, "0;(0)")) & " where IdFactura=" & TxtCampos(0).Text, CnnPrincipal, adOpenDynamic, adLockOptimistic
      AccionTool 17
    Else
      MsgBox "No se le pueden agregar remisiones a la factura si esta ya esta impresa o anulada", vbCritical, "No se pueden agregar remisiones"
    End If
  Else
    MsgBox "Debe seleccionar una planilla", vbCritical
  End If
End Sub



Private Sub CmdLiberarGuias_Click()
  If CpPermisoEspecial(16, CodUsuarioActivo, CnnPrincipal) = True Then
    AbrirRecorset rstUniversal, "SELECT IdFactura, Estado FROM facturas WHERE IdFactura = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstUniversal.RecordCount > 0 Then
      If rstUniversal!Estado = "I" Then
        If MsgBox("Esta seguro de liberar las guias para volver a facturar?", vbQuestion + vbYesNo, "Liberar guias de la factura") = vbYes Then
          AbrirRecorset rstUniversal, "UPDATE guias SET IdFactura=0, Facturada=0, IdPlanillaFactura=null WHERE IdFactura=" & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
        End If
      Else
        MsgBox "Solo se pueden liberar las guias de facturas impresas", vbInformation
      End If
    End If
  Else
    MsgBox "El usuario no tiene permiso para liberar guias"
  End If
End Sub

Private Sub CmdManPlanillas_Click()
  AbrirRecorset rstUniversal, "Select IdFactura, Estado, NroConceptos, NroGuias from facturas where Estado='D' and NroConceptos=0 and NroGuias=0 and IdFactura=" & Val(TxtCampos(0)), CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.EOF = False Then
    FraPlanillas.Height = 4035
    FraPlanillas.Enabled = True
    CmdManPlanillas.Visible = False
    VerPlanillas
  Else
    MsgBox "La factura debe esta digitada, no puede tener guias asiganadas ni conceptos", vbCritical
  End If
  CerrarRecorset rstUniversal
End Sub

Private Sub CmdMantenimientoConceptos_Click()
  AbrirRecorset rstUniversal, "Select IdFactura, Estado, NroConceptos, NroGuias from facturas where Estado='D' and NroGuias=0 and NroPlanillas=0 and IdFactura=" & Val(TxtCampos(0)), CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.EOF = False Then
    FraConceptos.Height = 3720
    FraConceptos.Enabled = True
    CmdMantenimientoConceptos.Visible = False
    CboConceptos.SetFocus
    VerConceptos
  Else
    MsgBox "La factura debe esta digitada, no puede tener guias asiganadas ni planillas", vbCritical
  End If
  CerrarRecorset rstUniversal
End Sub

Private Sub CmdQuitar_Click()
  II = 1
  Do While II <= LstConceptos.ListItems.Count
    If LstConceptos.ListItems(II).Checked = True Then
      AbrirRecorset rstUniversal, "Delete from conceptosfacturas where IdConceptoFactura=" & Val(LstConceptos.ListItems(II)) & " and IdFactura=" & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
      
      TxtCampos(8).Text = Val(Format(TxtCampos(8).Text, "0;(0)")) - Val(Format(LstConceptos.ListItems(II).SubItems(2), "0;(0)"))
      TxtCampos(19) = Val(TxtCampos(19)) - 1
      Liquidar
      AbrirRecorset rstUniversal, "Update Facturas set NroConceptos=NroConceptos-1, DctoComercial=" & Val(Format(TxtCampos(9).Text, "0;(0)")) & ", DctoFinanciero=" & Val(Format(TxtCampos(11).Text, "0;(0)")) & ", TotalFactura=" & Val(Format(TxtCampos(15).Text, "0;(0)")) & ", TOtros=" & Val(Format(TxtCampos(8).Text, "0;(0)")) & ", Saldo=" & Val(Format(TxtCampos(16).Text, "0;(0)")) & " where IdFactura=" & TxtCampos(0).Text, CnnPrincipal, adOpenDynamic, adLockOptimistic
      LstConceptos.ListItems.Remove (II)
    Else
      II = II + 1
    End If
  Loop
  AccionTool 17
End Sub

Private Sub CmdQuitarPlanilla_Click()
  Dim NroGuias As Integer
  II = 1
  
  Do While II <= LstPlanillas.ListItems.Count
    If LstPlanillas.ListItems(II).Checked = True Then
        AbrirRecorset rstUniversal, "select Guia from guias where IdFactura=" & Val(TxtCampos(0).Text) & " and IdPlanillaFactura=" & Val(LstPlanillas.ListItems(II)), CnnPrincipal, adOpenDynamic, adLockOptimistic
          NroGuias = Val(rstUniversal.RecordCount)
        CerrarRecorset rstUniversal
        If NroGuias = 0 Then
          AbrirRecorset rstUniversal, "Update facturas set NroPlanillas=NroPlanillas-1 where IdFactura=" & Val(TxtCampos(0)), CnnPrincipal, adOpenDynamic, adLockOptimistic
          AbrirRecorset rstUniversal, "Delete from facturasplanillas where IdPlanilla=" & Val(LstPlanillas.ListItems(II)), CnnPrincipal, adOpenDynamic, adLockOptimistic
        
          LstPlanillas.ListItems.Remove (II)
          TxtCampos(18) = Val(TxtCampos(18)) - 1
        Else
          MsgBox "Para eliminar la planilla " & LstPlanillas.SelectedItem & " debe quitarle primero las " & NroGuias & " guias que tiene asociadas", vbCritical
          II = II + 1
        End If
    Else
      II = II + 1
    End If
  Loop
  AccionTool 17
End Sub



Private Sub CmdVerConceptos_Click()
  VerConceptos
End Sub

Private Sub CmdVerGuias_Click()
  If rstTem.State = adStateOpen Then rstTem.Close
  rstTem.Open "SELECT guias.Guia, Ciudades.NmCiudad, FhEntradaBodega, Unidades, VrFlete, VrManejo, DocCliente, NmDestinatario, VrDeclarado, KilosFacturados " & _
              "FROM guias LEFT JOIN Ciudades ON guias.IdCiuDestino=Ciudades.IdCiudad " & _
              "WHERE IdFactura=" & Val(TxtCampos(0).Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Set GrillaFacturadas.DataSource = rstTem
  GrillaFacturadas.Tag = "Llena"
End Sub
Private Sub CmdVerPlanillas_Click()
  VerPlanillas
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  LaTecla (KeyCode), ToolFacturas
End Sub

Private Sub Form_Load()
  IconosTool ToolFacturas, Principal.IgListTool
  rstFacturas.CursorLocation = adUseServer
  rstTem.CursorLocation = adUseClient
  strSqlFacturas = "SELECT facturas.*, " & _
                "terceros.RazonSocial, formas_pago.NmFormaPago, centrosoperaciones.NmPuntoOperaciones " & _
                "FROM facturas " & _
                "LEFT JOIN terceros ON facturas.IdCliente = terceros.IDTercero " & _
                "LEFT JOIN formas_pago ON facturas.IdFormaPago = formas_pago.IdFormaPago " & _
                "LEFT JOIN centrosoperaciones ON facturas.codigo_centro_operaciones_fk = centrosoperaciones.IDPO "

  AbrirRecorset rstFacturas, strSqlFacturas & " Order by IdFactura Desc Limit 100", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Formatos rstFacturas
  Asignar rstFacturas
  Editando = False
  Dim rstTipoFacturas As New ADODB.Recordset
  rstTipoFacturas.CursorLocation = adUseClient
End Sub
Private Sub Asignar(rstAsignar As ADODB.Recordset)
  For II = 0 To 22
    TxtCampos(II).Text = rstAsignar.Fields(II) & ""
  Next
  TxtNmCliente = rstAsignar!RazonSocial
  LblFormaPago.Caption = rstAsignar!NmFormaPago & ""
  LblNmCentroOperaciones.Caption = rstAsignar!NmPuntoOperaciones
  
  If GrillaFacturadas.Tag = "Llena" Then
    Set rstTem.DataSource = Nothing
    Set GrillaFacturadas.DataSource = Nothing
    GrillaFacturadas.Tag = "Vacia"
  End If
  If LstConceptos.Tag = "Llena" Then
    LstConceptos.ListItems.Clear
    LstConceptos.Tag = "Vacia"
  End If
  If LstPlanillas.Tag = "Llena" Then
    LstPlanillas.ListItems.Clear
  End If

  If FraPlanillas.Enabled = True Then
    FraPlanillas.Enabled = False
    FraPlanillas.Height = 3675
    CmdManPlanillas.Visible = True
  End If
  
  If FraConceptos.Enabled = True Then
    FraConceptos.Enabled = False
    FraConceptos.Height = 3000
    CmdMantenimientoConceptos.Visible = True
  End If
  
End Sub
Private Sub Formatos(rstForma As ADODB.Recordset)
  For II = 0 To 22
    Set rstForma.Fields(II).DataFormat = TxtCampos(II).DataFormat
  Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstTem = Nothing
End Sub




Private Sub ToolFacturas_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub
Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
      If CpPermiso(3, CodUsuarioActivo, 2, CnnPrincipal) = True Then
        Desbloquear
        Limpiar
        If GrillaFacturadas.Tag = "Llena" Then
          Set rstTem.DataSource = Nothing
          Set GrillaFacturadas.DataSource = Nothing
          GrillaFacturadas.Tag = "Vacia"
        End If
        TxtCampos(1) = Format(Date, "dd/mm/yy") & " " & Format(Time, "h:m:s")
        TxtCampos(2) = Format(Date, "dd mmm yy")
        TxtCampos(3).SetFocus
        TxtCampos(4).Text = "D"
        TxtCampos(20).Text = "0"
        TxtCampos(22).Text = Coperaciones
      End If
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            Dim FechaFactura As Date
            FechaFactura = TxtCampos(1).Text
            Bloquear
            AbrirRecorset rstUniversal, "Update Facturas set IdCliente='" & TxtCampos(3).Text & "', Notas='" & TxtCampos(5) & "', FhVenceFac='" & Format(FechaFactura + Val(TxtCampos(20).Text), "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', IdFormaPago = " & Val(TxtCampos(21).Text) & ", codigo_centro_operaciones_fk = " & Val(TxtCampos(22).Text) & ", Plazo = " & Val(TxtCampos(20).Text) & " WHERE IdFactura=" & Val(TxtCampos(0)), CnnPrincipal, adOpenDynamic, adLockOptimistic
            Editando = False
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
            TxtCampos(0).Text = SacarConsecutivo("PreFactura", CnnPrincipal)
            AbrirRecorset rstUniversal, "INSERT INTO Facturas (IdFactura, FhFac, FhVenceFac, IdCliente, Estado, Notas, TFlete, TManejo, TOtros, DctoComercial, BaseCCial, DctoFinanciero, BaseFin, AntesDeDcto, Abonos, TotalFactura, Saldo, NroGuias, NroPlanillas, NroConceptos, Exportada, IdFormaPago, codigo_centro_operaciones_fk, Plazo) " & _
                                        "VALUES (" & Val(TxtCampos(0).Text) & ", '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', '" & Format(Date + Val(TxtCampos(20).Text), "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', '" & TxtCampos(3).Text & "', 'D', '" & TxtCampos(5).Text & "', 0, 0, 0, 0, " & Val(Format(TxtCampos(10).Text, "0;(0)")) & ", 0," & Val(Format(TxtCampos(12).Text, "0;(0)")) & ", " & Val(Format(TxtCampos(13).Text, "0;(0)")) & ", 0, 0, 0, 0, 0, 0, 0, " & Val(TxtCampos(21).Text) & ", " & Val(TxtCampos(22).Text) & ", " & Val(TxtCampos(20).Text) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
            Bloquear
            AccionTool 17
            AccionTool 11
            'If FufuLong = 6 Then CmdAgregarQuitar_Click
        End If
      End If
    Case 5  'Editar
      If CpPermiso(3, CodUsuarioActivo, 3, CnnPrincipal) = True Then
        If ExRecorset("Select IdFactura, Estado from Facturas where Estado='D' and IdFactura= " & Val(TxtCampos(0))) = True Then
          Editando = True
          Desbloquear
        Else
          MsgBox "Solo se pueden editar facturas digitadas", vbCritical
        End If
      End If
    Case 6 'Eliminar
      If CpPermiso(3, CodUsuarioActivo, 4, CnnPrincipal) = True Then
        If MsgBox("¿Esta seguro de anular la factura?", vbQuestion + vbYesNo) = vbYes Then
          If ExRecorset("Select IdFactura, Estado from facturas where IdFactura=" & Val(TxtCampos(0)) & " and Estado='I'") = True Then
            Dim boolError As Boolean
            boolError = False
            AbrirRecorset rstUniversal, "Select IdCxC from cuentas_cobrar where NroDocumento = " & Val(TxtCampos(0)) & " AND TipoFactura = 1", CnnPrincipal, adOpenDynamic, adLockOptimistic
            If rstUniversal.RecordCount > 0 Then
                Dim IdCuentaCobrar As Long
                IdCuentaCobrar = rstUniversal.Fields("IdCxC")
                If ExRecorset("Select IdReciboDet from recibos_caja_det where codigo_cuenta_cobrar_fk=" & IdCuentaCobrar) = True Then
                  boolError = True
                End If
                If ExRecorset("Select IdNotaCreditoDet from notas_credito_det where IdCxC=" & IdCuentaCobrar) = True Then
                  boolError = True
                End If
            End If
            If boolError = False Then
              AbrirRecorset rstUniversal, "Update guias Set IdFactura=0, Facturada=0, IdPlanillaFactura=null where IdFactura=" & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
              AbrirRecorset rstUniversal, "Update Facturas Set Estado='A', TFlete=0, TManejo=0, TOtros=0, TotalFactura=0, Saldo=0, NroGuias=0, NroPlanillas=0, NroConceptos=0 where IdFactura=" & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
              AbrirRecorset rstUniversal, "UPDATE facturas_venta SET Total = 0, VrFlete = 0, VrManejo = 0, VrOtros = 0 WHERE TipoFactura =1 AND Numero = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
              AbrirRecorset rstUniversal, "UPDATE cuentas_cobrar SET Total = 0, Abono = 0, Saldo = 0, VrFlete = 0, VrManejo = 0 WHERE TipoFactura =1 AND NroDocumento = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
              TxtCampos(4).Text = "A"
              AccionTool 17
              MsgBox "Factura Anulada con exito", vbInformation
            Else
              MsgBox "La factura tiene un recibo de caja o nota credito que la esta afectando y no se puede anular"
            End If
          Else
            MsgBox "La factura debe estar impresa", vbCritical
          End If
        End If
      End If
      
    Case 7 'Cancelar
      If Editando = True Then
        AccionTool 4
      Else
        Asignar rstFacturas
        Bloquear
      End If
    Case 9  'Buscar
      If Principal.ToolConsultas1.AbrirDevDatos("Numero de factura", "Digite el numero de la factura que desea buscar", 3, 0) = True Then
        AbrirRecorset rstUniversal, strSqlFacturas & " WHERE IdFactura=" & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          Asignar rstUniversal
        Else
          MsgBox "No se encontraron facturas con este numero", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If
    Case 11 'Primero
      UPrimero rstFacturas
      Asignar rstFacturas
    Case 12 'Anterior
      UAnterior rstFacturas
      Asignar rstFacturas
    Case 13 'Siguiente
      USiguiente rstFacturas
      Asignar rstFacturas
    Case 14 'Ultimo
      UUltimo rstFacturas
      Asignar rstFacturas
    Case 16 'Cerrar
      Set rstFacturas = Nothing
      Unload Me
    Case 17 'Actualizar
      rstFacturas.Requery
      Formatos rstFacturas
    Case 18 'Imprimir
    Case 19
      If TxtNmCliente.Text = "" Then
        AbrirRecorset rstUniversal, "Select IdTercero, RazonSocial from Terceros where IdTercero='" & TxtCampos(3).Text & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          TxtNmCliente.Text = rstUniversal!RazonSocial & ""
        End If
        CerrarRecorset rstUniversal
      End If
    Case 20
  End Select
End Sub
Private Sub Desbloquear()
  BotTool 3, 17, ToolFacturas, True
  FraCliente.Enabled = True
  FraNotas.Enabled = True
  CmdAgregarQuitar.Enabled = False
  CmdVerGuias.Enabled = False
End Sub

Private Sub Bloquear()
  BotTool 3, 17, ToolFacturas, False
  FraCliente.Enabled = False
  FraNotas.Enabled = False
  CmdAgregarQuitar.Enabled = True
  CmdVerGuias.Enabled = True
End Sub

Private Sub Limpiar()
  For II = 0 To 22
    TxtCampos(II).Text = ""
  Next
  TxtNmCliente.Text = ""
  LblFormaPago.Caption = ""
  LblNmCentroOperaciones.Caption = ""
End Sub
Function Validacion() As Boolean
  If TxtCampos(3).Text <> "" And TxtCampos(0).Text <> "0" Then
    If TxtCampos(22).Text <> "" And Val(TxtCampos(22).Text) <> 0 Then
      Validacion = True
    Else
      Validacion = False: MsgTit "La factura debe tener un centro operaciones": TxtCampos(22).SetFocus
    End If
  Else
    Validacion = False: MsgTit "La factura debe tener un cliente": TxtCampos(3).SetFocus
  End If
End Function

Private Sub ToolFacturas_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim rstFactura As New ADODB.Recordset
rstFactura.CursorLocation = adUseClient
rstFactura.Open "SELECT IdFactura, NroGuias, NroPlanillas, NroConceptos, Estado FROM facturas WHERE IdFactura=" & Val(TxtCampos(0)), CnnPrincipal, adOpenDynamic, adLockOptimistic
If rstFactura.EOF = False Then
  Select Case ButtonMenu.Key
    Case "Imp1"
      If Val(rstFactura!NroGuias) <> 0 And Val(rstFactura!NroPlanillas) = 0 And Val(rstFactura!NroConceptos) = 0 Then
        If rstFactura!Estado = "D" Then
          EstadoImpresoFactura rstFactura!IdFactura
        End If
        GImprimirFactura Val(TxtCampos(0)), 1
        AccionTool 17
      Else
        MsgBox "Solo se pueden imprimir facturas digitadas y para esta opcion no pueden tener conceptos ni planillas", vbCritical
      End If
      
    Case "Imp2"
      If Val(rstFactura!NroPlanillas) <> 0 And Val(rstFactura!NroGuias) = 0 And Val(rstFactura!NroConceptos) = 0 Then
        If rstFactura!Estado = "D" Then
          EstadoImpresoFactura rstFactura!IdFactura
          AbrirRecorset rstUniversal, "Update facturasplanillas set IdFactura=" & Val(TxtCampos(0).Text) & " where IdFactura=" & rstFactura!IdFactura, CnnPrincipal, adOpenDynamic, adLockOptimistic
        End If
        GImprimirFactura Val(TxtCampos(0)), 2
        AccionTool 17
      Else
        MsgBox "Solo se pueden imprimir facturas digitadas y para esta opcion no pueden tener conceptos ni guias", vbCritical
      End If
      
    Case "Imp3"
      If Val(rstFactura!NroConceptos) <> 0 And Val(rstFactura!NroPlanillas) = 0 And Val(rstFactura!NroGuias) = 0 Then
        If rstFactura!Estado = "D" Then
          EstadoImpresoFactura rstFactura!IdFactura
          AbrirRecorset rstUniversal, "Update conceptosfacturas set IdFactura=" & Val(TxtCampos(0).Text) & " where IdFactura=" & rstFactura!IdFactura, CnnPrincipal, adOpenDynamic, adLockOptimistic
        End If
        GImprimirFactura Val(TxtCampos(0)), 3
        AccionTool 17
      Else
        MsgBox "Solo se pueden imprimir facturas digitadas y para esta opcion no pueden tener guias ni planillas", vbCritical
      End If
      
    Case "Imp4"
      If Val(rstFactura!NroGuias) <> 0 And Val(rstFactura!NroPlanillas) = 0 And Val(rstFactura!NroConceptos) = 0 Then
        If rstFactura!Estado = "D" Then
          EstadoImpresoFactura rstFactura!IdFactura
        End If
        
        Mostrar_Reporte CnnPrincipal, 26, "SELECT * FROM sql_if_imp_factura WHERE IdFactura = " & Val(TxtCampos(0).Text), "Imprimir Factura", 2
        AccionTool 17
      Else
        MsgBox "Solo se pueden imprimir facturas digitadas y para esta opcion no pueden tener conceptos ni planillas", vbCritical
      End If
    
    Case "Imp5"
      If Val(rstFactura!NroPlanillas) <> 0 And Val(rstFactura!NroGuias) = 0 And Val(rstFactura!NroConceptos) = 0 Then
        If rstFactura!Estado = "D" Then
          EstadoImpresoFactura rstFactura!IdFactura
          AbrirRecorset rstUniversal, "Update facturasplanillas set IdFactura=" & Val(TxtCampos(0).Text) & " where IdFactura=" & rstFactura!IdFactura, CnnPrincipal, adOpenDynamic, adLockOptimistic
        End If
        Mostrar_Reporte CnnPrincipal, 35, "SELECT * FROM sql_if_imp_factura_planillas WHERE IdFactura = " & Val(TxtCampos(0).Text), "Imprimir Factura", 2
        AccionTool 17
      Else
        MsgBox "Solo se pueden imprimir facturas digitadas y para esta opcion no pueden tener conceptos ni guias", vbCritical
      End If
    
    Case "Imp6"
      If Val(rstFactura!NroConceptos) <> 0 And Val(rstFactura!NroPlanillas) = 0 And Val(rstFactura!NroGuias) = 0 Then
        If rstFactura!Estado = "D" Then
          EstadoImpresoFactura rstFactura!IdFactura
          AbrirRecorset rstUniversal, "Update conceptosfacturas set IdFactura=" & Val(TxtCampos(0).Text) & " where IdFactura=" & rstFactura!IdFactura, CnnPrincipal, adOpenDynamic, adLockOptimistic
        End If
        Mostrar_Reporte CnnPrincipal, 40, "SELECT * FROM sql_if_imp_factura_conceptos WHERE IdFactura = " & Val(TxtCampos(0).Text), "Imprimir Factura", 2
        AccionTool 17
      Else
        MsgBox "Solo se pueden imprimir facturas digitadas y para esta opcion no pueden tener guias ni planillas", vbCritical
      End If
    
    Case "Otr1"
      FrmCorregirGuias.Show 1
  End Select
End If
rstFactura.Close
Set rstFactura = Nothing
End Sub

Private Sub TxtCampos_Change(Index As Integer)
  If Index = 4 Then TxtEstado.Text = DevEstado(TxtCampos(4).Text)
End Sub

Private Sub TxtCampos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Select Case Index
    Case 3
      If KeyCode = vbKeyF2 Then
        Principal.ToolConsultas1.AbrirDevConsulta 7, CnnPrincipal
        TxtCampos(3).Text = Principal.ToolConsultas1.DatSt
      End If
    Case 22
      If KeyCode = vbKeyF2 Then
        FrmBuscarCO.Show 1
        TxtCampos(22).Text = FufuLo
      End If
  End Select
  
End Sub

Private Sub TxtCampos_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
    Case 5
      If KeyAscii = 13 Then
        KeyAscii = 0
      End If
  End Select
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtCampos_LostFocus(Index As Integer)
  Select Case Index
    Case 3
      If TxtCampos(3).Text = "0" Then
          TxtNmCliente = ""
          TxtCampos(3).Text = ""
          TxtCampos(10).Text = "0"
          TxtCampos(12).Text = "0"
          TxtCampos(13).Text = "0"
      Else
        AbrirRecorset rstUniversal, "SELECT IdTercero, RazonSocial, Plazo, terceros.IdFormaPago, NmFormaPago, terceros.IdFormaPago " & _
                                    "FROM terceros " & _
                                    "LEFT JOIN formas_pago ON terceros.IdFormaPago = formas_pago.IdFormaPago " & _
                                    "WHERE IdTercero='" & TxtCampos(3).Text & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          TxtNmCliente = rstUniversal!RazonSocial & ""
          TxtCampos(20).Text = rstUniversal.Fields("Plazo")
          LblFormaPago.Caption = rstUniversal.Fields("NmFormaPago")
          TxtCampos(21).Text = rstUniversal.Fields("IdFormaPago")
          TxtCampos(9).Text = 0
          TxtCampos(10).Text = 0
          TxtCampos(11).Text = 0
          TxtCampos(12).Text = 0
          TxtCampos(13).Text = 0
        Else
          TxtNmCliente = "": TxtCampos(3).Text = ""
        End If
        CerrarRecorset rstUniversal
      End If
    Case 22
      If Val(TxtCampos(22).Text) <> 0 Then
        AbrirRecorset rstUniversal, "SELECT IDPO, NmPuntoOperaciones From centrosoperaciones where IDPO=" & TxtCampos(22), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          LblNmCentroOperaciones.Caption = rstUniversal!NmPuntoOperaciones & ""
        Else
          LblNmCentroOperaciones.Caption = "": TxtCampos(22) = ""
        End If
        CerrarRecorset rstUniversal
      Else
          LblNmCentroOperaciones.Caption = "": TxtCampos(22) = ""
      End If
  End Select
End Sub

Private Sub Liquidar()
  Dim TFlete As Currency, TManejo As Currency, TOtros As Currency, DctoComercial As Currency, BaseCcial As Currency, BaseFin As Currency, TotalFactura As Currency, Abono As Currency
  TFlete = Val(Format(TxtCampos(6).Text, "0;(0)"))
  TManejo = Val(Format(TxtCampos(7).Text, "0;(0)"))
  TOtros = Val(Format(TxtCampos(8).Text, "0;(0)"))
  BaseCcial = Val(Format(TxtCampos(10).Text, "0;(0)"))
  BaseFin = Val(Format(TxtCampos(12).Text, "0;(0)"))
  TotalFactura = Format(TFlete + TManejo + TOtros, "0;(0)")
  TxtCampos(9).Text = Format(TotalFactura * BaseCcial / 100, "#,##0.00;(#,##0.00)")
  DctoComercial = Val(Format(TxtCampos(9).Text, "0;(0)"))
  TxtCampos(15).Text = Format(TotalFactura - DctoComercial, "#,##0.00;(#,##0.00)")
  TotalFactura = Val(Format(TxtCampos(15).Text, "0;(0)"))
  TxtCampos(11).Text = Format(TotalFactura * BaseFin / 100, "#,##0.00;(#,##0.00)")
  TxtCampos(16).Text = Format(TotalFactura - Abono, "#,##0.00;(#,##0.00)")
End Sub

Private Sub VerPlanillas()
  LstPlanillas.ListItems.Clear
  AbrirRecorset rstUniversal, "Select*from FacturasPlanillas Where IdFactura=" & Val(TxtCampos(0)), CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstUniversal.RecordCount > 0 Then
      Do While rstUniversal.EOF = False
        Set Item = LstPlanillas.ListItems.Add(, , rstUniversal!IdPlanilla)
          Item.SubItems(1) = rstUniversal!RelCliente
          Item.SubItems(2) = Format(rstUniversal!VrFletePlanilla, "#,##0.00;(#,##0.00)")
          Item.SubItems(3) = Format(rstUniversal!VrManejoPlanilla, "#,##0.00;(#,##0.00)")
          Item.SubItems(4) = Format(rstUniversal!NroGuiasPlanilla, "#,##0.00;(#,##0.00)")
          rstUniversal.MoveNext
      Loop
      LstPlanillas.Tag = "Llena"
    End If
  CerrarRecorset rstUniversal
End Sub

Private Sub VerConceptos()
  LstConceptos.ListItems.Clear
  AbrirRecorset rstUniversal, "Select ConceptosFacturas.*, conceptoscontables.NmConcepto from ConceptosFacturas, ConceptosContables Where (ConceptosFacturas.IdConcepto=ConceptosContables.IdConcepto) and IdFactura=" & Val(TxtCampos(0)), CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstUniversal.RecordCount > 0 Then
      Do While rstUniversal.EOF = False
        Set Item = LstConceptos.ListItems.Add(, , rstUniversal!IdConceptoFactura)
          Item.SubItems(1) = rstUniversal!NmConcepto
          Item.SubItems(2) = Format(rstUniversal!Valor, "#,##0.00;(#,##0.00)")
          rstUniversal.MoveNext
      Loop
      LstConceptos.Tag = "Llena"
    End If
  CerrarRecorset rstUniversal
End Sub

Private Sub TxtValorConcepto_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub EstadoImpresoFactura(intNumeroFactura As Long)
  Dim rstFactura As New ADODB.Recordset
  Dim NroFactura As Long
  Dim Plazo As Integer
  rstFactura.CursorLocation = adUseClient
  AbrirRecorset rstFactura, "SELECT facturas.Plazo, FhFac, TFlete, TManejo, TOtros, IdAsesor FROM facturas LEFT JOIN terceros ON facturas.IdCliente = terceros.IdTercero WHERE IdFactura = " & intNumeroFactura, CnnPrincipal, adOpenDynamic, adLockOptimistic
  Plazo = Val(rstFactura!Plazo)
  NroFactura = SacarConsecutivo("Facturas", CnnPrincipal)
  AbrirRecorset rstUniversal, "UPDATE guias SET IdFactura=" & NroFactura & " WHERE IdFactura=" & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
  Dim ValorLetras As String
  ValorLetras = UCase(CovLetras(Val(rstFactura.Fields("TFlete")) + Val(rstFactura.Fields("TManejo")) + Val(rstFactura.Fields("TOtros"))))
  AbrirRecorset rstUniversal, "UPDATE facturas set IdFactura=" & NroFactura & ", FhFac='" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', FhVenceFac='" & Format(Date + Plazo, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', Estado='I', ValorEnLetras='" & ValorLetras & "' WHERE IdFactura=" & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
  
  AbrirRecorset rstFactura, "SELECT IdFactura, FhFac, FhVenceFac, facturas.IdCliente, facturas.Plazo, TotalFactura, TFlete, TManejo, TOtros, codigo_centro_operaciones_fk, terceros.IdAsesor FROM facturas LEFT JOIN terceros ON facturas.IdCliente = terceros.IDTercero WHERE IdFactura = " & NroFactura, CnnPrincipal, adOpenDynamic, adLockOptimistic
  AbrirRecorset rstUniversal, "INSERT INTO facturas_venta (Numero, TipoFactura, Fecha, FhVence, IdTercero, Plazo, Total, VrFlete, VrManejo, VrOtros, IdPO, IdAsesor) VALUES (" & NroFactura & ", 1, '" & Format(rstFactura!FhFac, "yyyy/mm/dd") & "', '" & Format(rstFactura!FhVenceFac, "yyyy/mm/dd") & "', '" & rstFactura!IdCliente & "', " & rstFactura!Plazo & ", " & rstFactura!TotalFactura & ", " & rstFactura!TFlete & ", " & rstFactura!TManejo & ", 0, " & rstFactura!codigo_centro_operaciones_fk & ", " & rstFactura!IdAsesor & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
  AbrirRecorset rstUniversal, "INSERT INTO cuentas_cobrar(NroDocumento, TipoFactura, FechaDoc, FhVence, IdTercero, Total, Saldo, VrFlete, VrManejo, GuiaFactura, Condicion, IdAsesor, IdPO) VALUES (" & NroFactura & ", " & 1 & ", '" & Format(rstFactura!FhFac, "yyyy/mm/dd") & "', '" & Format(rstFactura!FhVenceFac, "yyyy/mm/dd") & "', '" & rstFactura!IdCliente & "', " & rstFactura!TotalFactura & ", " & rstFactura!TotalFactura & ", " & rstFactura!TFlete & ", " & rstFactura!TManejo & ", 0," & rstFactura!Plazo & "," & rstFactura!IdAsesor & "," & rstFactura!codigo_centro_operaciones_fk & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic

  TxtCampos(0).Text = NroFactura
End Sub
