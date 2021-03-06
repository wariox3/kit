VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmConsultaConductores 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta [Conductores]"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ChkInactivos 
      Caption         =   "Mostrar inactivos"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   1575
   End
   Begin VB.ComboBox CboCriterio 
      Height          =   315
      ItemData        =   "FrmConsultaConductores.frx":0000
      Left            =   120
      List            =   "FrmConsultaConductores.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox TxtBuscamos 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
   Begin MSDataGridLib.DataGrid GrillaDBasicos 
      Height          =   3615
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6376
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   6
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "IdConductor"
         Caption         =   "ID"
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
      BeginProperty Column01 
         DataField       =   "NmConductor"
         Caption         =   "Nombre"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   5894.929
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmConsultaConductores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstTem As New ADODB.Recordset
Private Sub CmdAceptar_Click()
  If rstTem.State <> adStateClosed Then If rstTem.EOF = False Then ElSt = rstTem.Fields("IdConductor") & ""
  Unload Me
End Sub
Private Sub CmdCancelar_Click()
  Unload Me
  ElSt = ""
End Sub
Private Sub Form_Load()
  CboCriterio.ListIndex = 1
End Sub
Sub Filtrar()
  Dim Consulta As String
  Consulta = "SELECT IdConductor, Concat(Nombre, ' ', Apellido1, ' ', Apellido2) as NmConductor FROM Conductores WHERE 1"
  
  If ChkInactivos.Value = 0 Then
    Consulta = Consulta & " AND ConductorInactivo = 0"
  End If
  
  If TxtBuscamos <> "" And CboCriterio.Text <> "" Then
    If CboCriterio.Text = "Codigo" Then
        Consulta = Consulta + " AND IdConductor like '%" & TxtBuscamos.Text & "%' order by NmConductor"
    End If
    If CboCriterio.Text = "Nombre" Then
        Consulta = Consulta + " AND Nombre like '%" & TxtBuscamos.Text & "%' order by NmConductor"
    End If
  End If
  AbrirRecorset rstTem, Consulta, Cn, adOpenStatic, adLockReadOnly
  Set GrillaDBasicos.DataSource = rstTem
End Sub
Private Sub GrillaDBasicos_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then CmdAceptar_Click
End Sub
Private Sub TxtBuscamos_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then GrillaDBasicos.SetFocus
End Sub
Private Sub TxtBuscamos_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Filtrar
End Sub

