VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmInfoCuenta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informacion de la cuenta..."
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Aceptar / Salir"
      Height          =   255
      Left            =   9720
      TabIndex        =   54
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   1380
      Left            =   120
      TabIndex        =   61
      Top             =   120
      Width           =   5415
      Begin VB.CheckBox ChkInactivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Inactivo"
         Height          =   255
         Left            =   4320
         TabIndex        =   62
         Top             =   960
         Width           =   975
      End
      Begin MSMask.MaskEdBox TxtFechaIng 
         Height          =   255
         Left            =   3960
         TabIndex        =   63
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtIdCliente 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   64
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label LblCLientes 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Telefono:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   71
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   70
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label LblCLientes 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ingreso:"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   68
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label LblCLientes 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nit:"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   67
         Top             =   240
         Width           =   240
      End
      Begin VB.Label LblCLientes 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   2
         Left            =   255
         TabIndex        =   66
         Top             =   600
         Width           =   600
      End
      Begin VB.Label LblNmCliente 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   65
         Top             =   600
         Width           =   4335
      End
   End
   Begin VB.Frame FraRecogidas 
      Caption         =   "Programacion de recogidas"
      Enabled         =   0   'False
      Height          =   2100
      Left            =   5640
      TabIndex        =   50
      Top             =   120
      Width           =   5775
      Begin MSComCtl2.DTPicker DPHoraRec 
         Height          =   255
         Left            =   1200
         TabIndex        =   69
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Format          =   49807361
         CurrentDate     =   38971
      End
      Begin VB.CheckBox ChKRecoge 
         Caption         =   "Recoger"
         Height          =   255
         Left            =   1200
         TabIndex        =   55
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtDirRecogidas 
         Height          =   285
         Left            =   1200
         TabIndex        =   53
         Top             =   960
         Width           =   4455
      End
      Begin VB.TextBox TxtRuta 
         Height          =   285
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   52
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox TxtEncargadoRec 
         Height          =   285
         Left            =   1200
         TabIndex        =   51
         Top             =   1680
         Width           =   4455
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "H Recogidas:"
         Height          =   195
         Index           =   45
         Left            =   120
         TabIndex        =   60
         Top             =   615
         Width           =   975
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "Dir Recogidas:"
         Height          =   195
         Index           =   3
         Left            =   45
         TabIndex        =   59
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ruta de rec:"
         Height          =   195
         Left            =   210
         TabIndex        =   58
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label LblConsulta 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   57
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "Encargado:"
         Height          =   195
         Index           =   7
         Left            =   270
         TabIndex        =   56
         Top             =   1680
         Width           =   825
      End
   End
   Begin VB.Frame FraTpCtaFlete 
      Caption         =   "Tipo de Cuenta Flete"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   6000
      TabIndex        =   46
      Top             =   4800
      Width           =   1935
      Begin VB.CheckBox ChkTpCtaFleCor 
         Caption         =   "Cuenta Corriente"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox ChkTpCtaFleCon 
         Caption         =   "Contado"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox ChkTpCtaConFleEnt 
         Caption         =   "Contra Entrega"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame FraTpCtaMan 
      Caption         =   "Tipo de Cuenta Manejo"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   8880
      TabIndex        =   42
      Top             =   4800
      Width           =   1935
      Begin VB.CheckBox ChkTpCtaManCtaCor 
         Caption         =   "Cuenta Corriente"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox ChkTpCtaManCon 
         Caption         =   "Contado"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox ChkTpCtaManCtraEnt 
         Caption         =   "Contra Entrega"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame FraServicios 
      Caption         =   "Servicios"
      Enabled         =   0   'False
      Height          =   1935
      Left            =   120
      TabIndex        =   40
      Top             =   4440
      Width           =   2535
      Begin MSComctlLib.ListView LstServicios 
         Height          =   1575
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2778
         View            =   2
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
         NumItems        =   0
      End
   End
   Begin VB.Frame FraCuentas 
      Caption         =   "Condiciones comerciales"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   120
      TabIndex        =   27
      Top             =   2280
      Width           =   3495
      Begin MSMask.MaskEdBox TxtDctoPieFac 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0,000E+00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   28
         Top             =   600
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDctoFinanciero 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   29
         Top             =   960
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   5
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtAntesDe 
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
         Left            =   2640
         TabIndex        =   30
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPlazo 
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
         Left            =   2640
         TabIndex        =   31
         Top             =   960
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtCupoCredito 
         Height          =   255
         Left            =   1200
         TabIndex        =   32
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "Dcto Fin:"
         Height          =   195
         Index           =   16
         Left            =   435
         TabIndex        =   39
         Top             =   960
         Width           =   645
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "Dcto Pie Fac:"
         Height          =   195
         Index           =   15
         Left            =   105
         TabIndex        =   38
         Top             =   600
         Width           =   975
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "Plazo:"
         Height          =   195
         Index           =   18
         Left            =   2040
         TabIndex        =   37
         Top             =   960
         Width           =   435
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "Ant de:"
         Height          =   195
         Index           =   17
         Left            =   1965
         TabIndex        =   36
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   1800
         TabIndex        =   35
         Top             =   960
         Width           =   120
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   34
         Left            =   1800
         TabIndex        =   34
         Top             =   600
         Width           =   120
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "Cupo Credito:"
         Height          =   195
         Index           =   44
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame FraManejo 
      Caption         =   "Seguro o manejo:"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   3480
      Begin MSMask.MaskEdBox TxtPorManejo 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   5
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtVrUnidad 
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
         Left            =   2160
         TabIndex        =   23
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   36
         Left            =   1080
         TabIndex        =   26
         Top             =   240
         Width           =   120
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "Man:"
         Height          =   195
         Index           =   35
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   360
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "Min x Uni:"
         Height          =   195
         Index           =   37
         Left            =   1320
         TabIndex        =   24
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame FraMas 
      Caption         =   "Carta porte"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   3360
      TabIndex        =   16
      Top             =   4800
      Width           =   1725
      Begin VB.OptionButton OptCp 
         Caption         =   "No"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton OptCp 
         Caption         =   "Si"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton OptCp 
         Caption         =   "Opcional"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "Devolver Carta Porte"
         Height          =   195
         Index           =   46
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1485
      End
   End
   Begin VB.Frame FraDescuentos 
      Caption         =   "Liquidacion"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   3720
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
      Begin VB.CheckBox ChkAdicional 
         Alignment       =   1  'Right Justify
         Caption         =   "Adicional"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.CheckBox ChkUnidad 
         Alignment       =   1  'Right Justify
         Caption         =   "Unidad"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox ChkKilo 
         Alignment       =   1  'Right Justify
         Caption         =   "Kilo"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
      Begin MSMask.MaskEdBox TxtDctoKilo 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   5
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDctoUni 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   5
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtMinimos 
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
         Left            =   480
         TabIndex        =   11
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   4
         Format          =   "0"
         PromptChar      =   "_"
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "Min:"
         Height          =   195
         Index           =   24
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   300
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   33
         Left            =   1560
         TabIndex        =   14
         Top             =   480
         Width           =   120
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   32
         Left            =   1560
         TabIndex        =   13
         Top             =   840
         Width           =   120
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "Dctos"
         Height          =   195
         Index           =   28
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   420
      End
   End
   Begin VB.Frame FraBas 
      Enabled         =   0   'False
      Height          =   2055
      Left            =   5640
      TabIndex        =   0
      Top             =   2280
      Width           =   5775
      Begin VB.TextBox TxtListaPrecioC 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   2
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox TxtObservaciones 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   1215
         Left            =   1200
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "FrmInfoCuenta.frx":0000
         Top             =   600
         Width           =   4485
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "Lista Precios:"
         Height          =   195
         Index           =   41
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   945
      End
      Begin VB.Label LblCLientes 
         AutoSize        =   -1  'True
         Caption         =   "Comentarios:"
         Height          =   195
         Index           =   42
         Left            =   150
         TabIndex        =   3
         Top             =   600
         Width           =   915
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   2760
      X2              =   11400
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      Index           =   0
      X1              =   11400
      X2              =   2760
      Y1              =   6360
      Y2              =   6360
   End
End
Attribute VB_Name = "FrmInfoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  AbrirRecorset rstUniversal, "SELECT*From Negociaciones where Id=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.EOF = False Then
    Set Item = LstServicios.ListItems.Add(1, , "Paqueteo")
    Set Item = LstServicios.ListItems.Add(2, , "Semi-Masivo")
    Set Item = LstServicios.ListItems.Add(3, , "Masivo")
    Set Item = LstServicios.ListItems.Add(4, , "Urbanos/Local")
    Set Item = LstServicios.ListItems.Add(5, , "Encomiendas")
    TxtIdCliente = rstUniversal!Id
    LblNmCliente = rstUniversal!NmNegociacion & ""
    TxtFechaIng = rstUniversal!FecIng
    ChkInactivo.value = DevCheck(rstUniversal.Fields("Inactivo"))
    ChkTpCtaFleCor.value = DevCheck(rstUniversal.Fields("TpCtaFleCor").value)
    ChkTpCtaFleCon.value = DevCheck(rstUniversal!TpCtaFleCon)
    ChkTpCtaConFleEnt.value = DevCheck(rstUniversal!TpCtaConFleEnt)
    ChkTpCtaManCtaCor.value = DevCheck(rstUniversal!TpCtaManCtaCor)
    ChkTpCtaManCon.value = DevCheck(rstUniversal!TpCtaManCon)
    ChkTpCtaManCtraEnt.value = DevCheck(rstUniversal!TpCtaManCtraEnt)
    TxtCupoCredito = rstUniversal!CupoCredito
    TxtDctoPieFac = rstUniversal!DctoPieFac & ""
    TxtAntesDe = rstUniversal!AntesDe & ""
    TxtDctoFinanciero = rstUniversal!DctoProPag & ""
    TxtPlazo = rstUniversal!Plazo & ""
    TxtPorManejo = rstUniversal!PorManejo
    TxtVrUnidad = rstUniversal!MinUniManejo
    ChkKilo.value = DevCheck(rstUniversal!ManKilo)
    TxtDctoKilo = rstUniversal!DctoK
    ChkUnidad.value = DevCheck(rstUniversal!ManUni)
    TxtDctoUni = rstUniversal!DctoU
    ChkAdicional.value = DevCheck(rstUniversal!ManAdicional)
    TxtMinimos = rstUniversal!Minimos
    DPHoraRec.value = rstUniversal!HorarioRecoge
    ChKRecoge.value = DevCheck(rstUniversal!Recoge)
    TxtDirRecogidas = rstUniversal!DirBodega & ""
    TxtRuta = rstUniversal!IdRutaRecogida & ""
    TxtEncargadoRec = rstUniversal!EncargadoRec & ""
    TxtListaPrecioC = rstUniversal!ListaPrecios
    TxtObservaciones = rstUniversal!Observaciones & ""
    OptCp(Val(rstUniversal!CartaPorte)).value = True
    LstServicios.ListItems(1).Checked = rstUniversal.Fields("ManPaqueteo").value
    LstServicios.ListItems(2).Checked = rstUniversal.Fields("ManSemiMasivo").value
    LstServicios.ListItems(3).Checked = rstUniversal.Fields("ManMasivo").value
    LstServicios.ListItems(4).Checked = rstUniversal.Fields("ManLocal").value
    LstServicios.ListItems(5).Checked = rstUniversal.Fields("ManEncomiendas").value
  End If
CerrarRecorset rstUniversal
End Sub
