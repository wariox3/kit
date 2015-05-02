VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmConsultaGeneral 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta..."
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtBuscamos 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin VB.ComboBox CboCriterio 
      Height          =   315
      ItemData        =   "FrmConsultaGeneral.frx":0000
      Left            =   120
      List            =   "FrmConsultaGeneral.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid Grilla 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6615
      _ExtentX        =   11668
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
         DataField       =   ""
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
         DataField       =   ""
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   5070.047
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmConsultaGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstTem As New ADODB.Recordset, II As Integer
Private Sub CboCriterio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub CmdAceptar_Click()
  If rstTem.State <> adStateClosed Then If rstTem.EOF = False Then Ello = rstTem.Fields(0)
  Unload Me
End Sub
Private Sub CmdCancelar_Click()
  Ello = 0
  Unload Me
End Sub
Private Sub Form_Load()
  For II = 1 To Len(ElSt)
    If Mid(ElSt, II, 1) = "," Then
      CboCriterio.AddItem TxtBuscamos.Text
      TxtBuscamos.Text = ""
    Else
      TxtBuscamos.Text = TxtBuscamos.Text & Mid(ElSt, II, 1)
    End If
  Next
  CboCriterio.Tag = TxtBuscamos.Text
  TxtBuscamos.Text = ""
  Me.Caption = Me.Caption & " [" & CboCriterio.Tag & "]"
  Grilla.Columns(0).DataField = CboCriterio.List(0)
  Grilla.Columns(1).DataField = CboCriterio.List(1)
  CboCriterio.ListIndex = 1
End Sub
Private Sub Filtrar()
  Dim Consulta As String
  Consulta = "SELECT " & CboCriterio.List(0) & ", " & CboCriterio.List(1) & " FROM " & CboCriterio.Tag
  If TxtBuscamos <> "" And CboCriterio.Text <> "" Then
    If CboCriterio.ListIndex = 0 Then Consulta = Consulta + " where " & CboCriterio.List(0) & " like '" & TxtBuscamos.Text & "%' order by " & CboCriterio.List(1)
    If CboCriterio.ListIndex = 1 Then Consulta = Consulta + " where " & CboCriterio.List(1) & " like '" & TxtBuscamos.Text & "%' order by " & CboCriterio.List(1)
  End If
  AbrirRecorset rstTem, Consulta, Cn, adOpenStatic, adLockReadOnly
  Set Grilla.DataSource = rstTem
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstTem = Nothing
End Sub

Private Sub Grilla_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then CmdAceptar_Click
End Sub
Private Sub TxtBuscamos_GotFocus()
  EnfocarT TxtBuscamos
End Sub
Private Sub TxtBuscamos_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then Grilla.SetFocus
End Sub
Private Sub TxtBuscamos_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Filtrar
End Sub
