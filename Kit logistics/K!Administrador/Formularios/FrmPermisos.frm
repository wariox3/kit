VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmPermisos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Permisos"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAgregarFormulario 
      Caption         =   "Agregar formulario"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   255
      Left            =   4800
      TabIndex        =   1
      Top             =   6000
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid GrillaPermisos 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "NmFormulario"
         Caption         =   "Formulario"
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
         DataField       =   "Ingreso"
         Caption         =   "Ingreso"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0.000E+00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Nuevo"
         Caption         =   "Nuevo"
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
         DataField       =   "Editar"
         Caption         =   "Editar"
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
         DataField       =   "Eliminar"
         Caption         =   "Eliminar"
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
            ColumnWidth     =   3404.977
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   705.26
         EndProperty
      EndProperty
   End
   Begin VB.Label LblNmUsuario 
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
      Left            =   600
      TabIndex        =   4
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label LblIdUsuario 
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
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "FrmPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstPermisos As New ADODB.Recordset

Private Sub CmdAgregarFormulario_Click()
  FufuLo = LblIdUsuario.Caption
  FrmAgregarFormulario.Show 1
  If rstPermisos.State = adStateOpen Then rstPermisos.Close
  rstPermisos.Open "select IdUsuario, permisos.IdFormulario, Ingreso, Nuevo, Editar, Eliminar, NmFormulario From (permisos left join formularios on((permisos.IdFormulario = formularios.IdFormulario))) where IdUsuario=" & Val(LblIdUsuario), CnnPrincipal, adOpenDynamic, adLockOptimistic
  Set GrillaPermisos.DataSource = rstPermisos
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  rstPermisos.CursorLocation = adUseClient
  LblIdUsuario.Caption = FufuLo
  LblNmUsuario.Caption = FufuSt
  rstPermisos.Open "select IdUsuario, permisos.IdFormulario, Ingreso, Nuevo, Editar, Eliminar, NmFormulario From (permisos left join formularios on((permisos.IdFormulario = formularios.IdFormulario))) where IdUsuario=" & Val(LblIdUsuario), CnnPrincipal, adOpenDynamic, adLockOptimistic
  Set GrillaPermisos.DataSource = rstPermisos
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstPermisos = Nothing
End Sub
