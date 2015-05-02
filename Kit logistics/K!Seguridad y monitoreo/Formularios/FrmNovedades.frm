VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmNovedades 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Novedades..."
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSolucionar 
      Caption         =   "Solucionar seleccionada"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Frame FraSolucion 
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   10215
      Begin VB.TextBox TxtSolucion 
         Height          =   1095
         Left            =   840
         TabIndex        =   15
         Top             =   240
         Width           =   9255
      End
      Begin VB.Label LblNotas 
         AutoSize        =   -1  'True
         Caption         =   "Solucion:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "..."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton CmdCancelarAccion 
      Caption         =   "&Cancelar accion"
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Frame FraNovedad 
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   10215
      Begin VB.TextBox TxtNovedad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   9015
      End
      Begin VB.TextBox TxtIdNovedad 
         Height          =   285
         Left            =   360
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TxtComentarios 
         Height          =   735
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   9015
      End
      Begin VB.Label LblIdNov 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   210
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Comentario:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   840
      End
   End
   Begin MSComCtl2.DTPicker DPicHora 
      Height          =   285
      Left            =   8640
      TabIndex        =   6
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      Format          =   20578306
      CurrentDate     =   39153
   End
   Begin MSComCtl2.DTPicker DPicFecha 
      Height          =   285
      Left            =   6360
      TabIndex        =   5
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      Format          =   20578305
      CurrentDate     =   39153
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "Agregar >>"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   255
      Left            =   8280
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin MSComctlLib.ImageList ImListLista 
      Left            =   0
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNovedades.frx":0000
            Key             =   "Ok"
            Object.Tag             =   "Ok"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNovedades.frx":0A14
            Key             =   "Pendiente"
            Object.Tag             =   "Pendiente"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LstNovedades 
      Height          =   1935
      Left            =   120
      TabIndex        =   21
      Top             =   360
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   3413
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImListLista"
      SmallIcons      =   "ImListLista"
      ColHdrIcons     =   "ImListLista"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "UsuIng"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fh/Hr Ingreso"
         Object.Width           =   2910
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fh/Hr Novedad"
         Object.Width           =   2910
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Novedad"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Comentarios"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "UsuSol"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Fh/Hr Sol"
         Object.Width           =   2910
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Solucion"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label LblMensaje 
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
      Height          =   255
      Left            =   3840
      TabIndex        =   20
      Top             =   120
      Width           =   6495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Monitoreo:"
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
      TabIndex        =   19
      Top             =   120
      Width           =   915
   End
   Begin VB.Label LblDespacho 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   1080
      TabIndex        =   18
      Top             =   120
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   195
      Left            =   5760
      TabIndex        =   12
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hora:"
      Height          =   195
      Left            =   8160
      TabIndex        =   11
      Top             =   4440
      Width           =   390
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   120
      X2              =   10320
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      Index           =   0
      X1              =   10320
      X2              =   120
      Y1              =   2760
      Y2              =   2760
   End
End
Attribute VB_Name = "FrmNovedades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdAgregar_Click()
  CmdAgregar.Enabled = False
  CmdSolucionar.Enabled = False
  FraNovedad.Visible = True
  Me.Height = 5100
  CmdGuardar.Caption = "Agregar novedad"
  TxtIdNovedad.SetFocus
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub CmdCancelarAccion_Click()
  CmdAgregar.Enabled = True
  CmdSolucionar.Enabled = True
  Me.Height = 3015
  CmdGuardar.Caption = "..."
  FraSolucion.Visible = False
  FraNovedad.Visible = False
  TxtIdNovedad.Text = ""
  TxtNovedad.Text = ""
  TxtComentarios.Text = ""
  TxtSolucion.Text = ""
End Sub
Private Sub CmdGuardar_Click()
  If CmdGuardar.Caption = "Agregar novedad" Then
    If TxtIdNovedad.Text <> "" Then
      rstUniversal.Open "INSERT INTO NovedadesMonitoreo (IdMonitoreo, IdNovedad, Comentarios, UsuIng, FhIngreso, FhNovedad, Solucion, UsuSol, FhSolucion, Solucionada) VALUES (" & Val(LblDespacho) & ", " & Val(TxtIdNovedad.Text) & ", '" & TxtComentarios.Text & "', " & CodUsuarioActivo & ", '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', '" & Format(DPicFecha.Value, "yyyy/mm/dd") & " " & Format(DPicHora.Value, "h:m:s") & "','', " & CodUsuarioActivo & ",'" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "',0)", CnnPrincipal, adOpenDynamic, adLockOptimistic
      rstUniversal.Open "Update MonitoreoVehiculos set EnNovedad=1 where ID=" & LblDespacho, CnnPrincipal, adOpenDynamic, adLockOptimistic
      TxtIdNovedad.Text = ""
      TxtNovedad.Text = ""
      TxtComentarios.Text = ""
      VerNovedades
      MsgBox "La novedad se a ingresado con exito", vbInformation
    Else
      MsgBox "Debe elegir una novedad para agragar", vbCritical, "Elija una novedad": TxtIdNovedad.SetFocus
    End If
  Else
    If LstNovedades.ListItems.Count > 0 Then
      If MsgBox("Va a solucionar la novedad con:" & Chr(13) & TxtSolucion.Text & Chr(13) & "¿Esta seguro que dese solucionar la novedad [" & LstNovedades.SelectedItem & "] ahora?", vbQuestion + vbYesNo) = vbYes Then
        rstUniversal.Open "Update NovedadesMonitoreo set Solucionada=1, Solucion='" & TxtSolucion & "', FHSolucion='" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', UsuSol='" & CodUsuarioActivo & "' where ID=" & LstNovedades.SelectedItem, CnnPrincipal, adOpenDynamic, adLockOptimistic
        TxtSolucion.Text = ""
        VerNovedades
        For II = 1 To LstNovedades.ListItems.Count
          If LstNovedades.ListItems(II).Icon = "Pendiente" Then Me.Tag = "N"
        Next
        If Me.Tag = "N" Then
          Me.Tag = ""
        Else
          rstUniversal.Open "Update MonitoreoVehiculos set EnNovedad=0 where ID=" & LblDespacho, CnnPrincipal, adOpenDynamic, adLockOptimistic
        End If
      End If
    Else
      MsgBox "Esta guia no tiene novedades", vbCritical, "Guia sin novedad"
    End If
  End If
  Me.Height = 3015
  CmdGuardar.Caption = "..."
  CmdAgregar.Enabled = True
  CmdSolucionar.Enabled = True
End Sub
Private Sub CmdSolucionar_Click()
  CmdSolucionar.Enabled = False
  CmdAgregar.Enabled = False
  FraSolucion.Visible = True
  Me.Height = 5100
  CmdGuardar.Caption = "Solucionar novedad"
  TxtSolucion.SetFocus
End Sub
Private Sub DPicFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub DPicHora_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub Form_Load()
FufuDo = 1
LblDespacho = FufuLo
DPicHora = Time
DPicFecha = Date
If FufuDo = 1 Then
  CmdAgregar.Enabled = True
  CmdSolucionar.Enabled = True
Else
  LblMensaje.Caption = "Solo puede ver novedades"
End If
  VerNovedades
End Sub

Sub VerNovedades()
  LstNovedades.ListItems.Clear
  rstUniversal.Open "SELECT NovedadesMonitoreo.*, CausalesNovedadMonitoreo.NmNovedad FROM NovedadesMonitoreo INNER JOIN CausalesNovedadMonitoreo ON NovedadesMonitoreo.IdNovedad = CausalesNovedadMonitoreo.IdNovedad Where IdMonitoreo=" & FufuLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    If Val(rstUniversal.Fields("Solucionada")) = 0 Then
      Set Item = LstNovedades.ListItems.Add(, , rstUniversal!ID, "Pendiente", "Pendiente")
    Else
      Set Item = LstNovedades.ListItems.Add(, , rstUniversal!ID, "Ok", "Ok")
    End If
    Item.SubItems(1) = rstUniversal!UsuIng & ""
    Item.SubItems(2) = Format(rstUniversal!FHIngreso, "dd/mmm/yy hh:mm:ss")
    Item.SubItems(3) = Format(rstUniversal!FHNovedad, "dd/mmm/yy hh:mm:ss")
    Item.SubItems(4) = rstUniversal!NmNovedad & ""
    Item.SubItems(5) = rstUniversal!Comentarios & ""
    Item.SubItems(6) = rstUniversal!UsuSol & ""
    Item.SubItems(7) = Format(rstUniversal!FHSolucion, "dd/mmm/yy hh:mm:ss")
    Item.SubItems(8) = rstUniversal!Solucion & ""
    rstUniversal.MoveNext
  Loop
  rstUniversal.Close
End Sub

Private Sub TxtIdNovedad_GotFocus()
  EnfocarT TxtIdNovedad
End Sub

Private Sub TxtIdNovedad_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirConsultaGral "IdNovedad", "NmNovedad", "CausalesNovedadMonitoreo", CnnPrincipal
    TxtIdNovedad.Text = Principal.ToolConsultas1.DatLo
  End If
End Sub

Private Sub TxtIdNovedad_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdNovedad_Validate(Cancel As Boolean)
  rstUniversal.Open "select IdNovedad, NmNovedad from CausalesNovedadMonitoreo where IdNovedad=" & Val(TxtIdNovedad), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = False Then
    TxtNovedad = rstUniversal!NmNovedad
  Else
    TxtIdNovedad.Text = "": TxtNovedad.Text = ""
  End If
  rstUniversal.Close
End Sub



