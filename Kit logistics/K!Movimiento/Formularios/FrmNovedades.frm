VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmNovedades 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nove dades..."
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdRegistro 
      Caption         =   "Registro"
      Height          =   375
      Left            =   4680
      TabIndex        =   18
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton CmdEnviarCorreo 
      Caption         =   "Enviar correo"
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton CmdImpirmir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   10560
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton CmdSolucionar 
      Caption         =   "Solucionar seleccionada"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "Agregar nueva"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   2175
   End
   Begin MSComctlLib.ListView LstNovedades 
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   4895
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
   Begin MSComctlLib.ImageList ImListLista 
      Left            =   0
      Top             =   3120
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
   Begin VB.Frame FraNovedad 
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Visible         =   0   'False
      Width           =   11775
      Begin VB.CommandButton CmdGuardar 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CommandButton CmdCancelarAccion 
         Caption         =   "&Cancelar "
         Height          =   375
         Left            =   2520
         TabIndex        =   21
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox TxtComentarios 
         Height          =   735
         Left            =   1080
         TabIndex        =   11
         Top             =   600
         Width           =   10575
      End
      Begin VB.TextBox TxtIdNovedad 
         Height          =   285
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TxtNovedad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   10575
      End
      Begin MSComCtl2.DTPicker DPicHora 
         Height          =   375
         Left            =   9240
         TabIndex        =   19
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49741825
         CurrentDate     =   38971
      End
      Begin MSComCtl2.DTPicker DPicFecha 
         Height          =   375
         Left            =   6720
         TabIndex        =   20
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49741825
         CurrentDate     =   38971
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         Height          =   195
         Left            =   8640
         TabIndex        =   24
         Top             =   1560
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   6120
         TabIndex        =   23
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Comentario:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   840
      End
      Begin VB.Label LblIdNov 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   210
      End
   End
   Begin MSComctlLib.ListView LstNovedadesMonitoreo 
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1931
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Guia:"
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
      TabIndex        =   17
      Top             =   120
      Width           =   465
   End
   Begin VB.Label LblIdMonitoreo 
      Height          =   255
      Left            =   9480
      TabIndex        =   15
      Top             =   0
      Width           =   735
   End
   Begin VB.Label LblIdDespacho 
      Height          =   255
      Left            =   8640
      TabIndex        =   14
      Top             =   0
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   120
      X2              =   11880
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   120
      X2              =   11880
      Y1              =   3720
      Y2              =   3720
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
      Left            =   6240
      TabIndex        =   5
      Top             =   45
      Width           =   2175
   End
   Begin VB.Label LblGuia 
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
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   75
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      Index           =   0
      X1              =   11880
      X2              =   120
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      Index           =   1
      X1              =   12000
      X2              =   120
      Y1              =   5040
      Y2              =   5040
   End
End
Attribute VB_Name = "FrmNovedades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents oMail As clsCDOmail
Attribute oMail.VB_VarHelpID = -1

Private Sub CmdAgregar_Click()
  CmdAgregar.Enabled = False
  CmdSolucionar.Enabled = False
  FraNovedad.Visible = True
  Me.Height = 7500
  TxtIdNovedad.SetFocus
End Sub

Private Sub CmdCancelarAccion_Click()
  CmdAgregar.Enabled = True
  CmdSolucionar.Enabled = True
  Me.Height = 5460
  FraNovedad.Visible = False
  TxtIdNovedad.Text = ""
  TxtNovedad.Text = ""
  TxtComentarios.Text = ""
End Sub

Private Sub CmdEnviarCorreo_Click()
  Dim Mensaje As String
  Dim MsgNovedad As String
  Dim email As String
  AbrirRecorset rstUniversal, "SELECT Email FROM guias LEFT JOIN terceros ON guias.Cuenta = terceros.IdTercero WHERE Guia = " & LblGuia.Caption, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.Fields("Email") & "" = "" Then
    MsgBox "El cliente no tiene un correo electronico valido", vbCritical, "Email no valido"
    Exit Sub
  Else
    email = rstUniversal.Fields("Email")
  End If
  CerrarRecorset rstUniversal
  AbrirRecorset rstUniversal, "SELECT guias.*, NmCiudad " & _
  "FROM guias " & _
  "LEFT JOIN ciudades ON guias.IdCiuDestino = ciudades.IdCiudad " & _
  "WHERE Guia = " & LblGuia.Caption, CnnPrincipal, adOpenDynamic, adLockOptimistic
  
  Mensaje = "Cordial saludo," & Chr(13) & _
  "Le informamos que se presento una novedad con:" & Chr(13) & _
  "Guia numero: " & LblGuia.Caption & Chr(13) & _
  "Documento o guia del cliente numero: " & rstUniversal.Fields("DocCliente") & "" & Chr(13) & _
  "Destinatario: " & rstUniversal.Fields("NmDestinatario") & "" & Chr(13) & _
  "Destino: " & rstUniversal.Fields("NmCiudad") & "" & Chr(13) & "_____________________________________" & Chr(13)
  
  CerrarRecorset rstUniversal
  
  AbrirRecorset rstUniversal, "SELECT novedades.*, causalesnovedad.NmNovedad " & _
                              "FROM novedades " & _
                              "LEFT JOIN causalesnovedad ON novedades.IdNovedad = causalesnovedad.IdNovedad " & _
                              "WHERE Guia=" & FufuLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    MsgNovedad = MsgNovedad & "Novedad: " & rstUniversal.Fields("NmNovedad") & Chr(13) & _
    "Comentario: " & rstUniversal.Fields("Comentarios") & Chr(13) & _
    "_____________________________________" & Chr(13)
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
  
  Mensaje = Mensaje & MsgNovedad & Chr(13)
  Mensaje = Mensaje & "En espera de sus comentarios e informar como proceder con esta novedad." & Chr(13) & Chr(13)
  Mensaje = Mensaje & "Cordialmente," & Chr(13)
  Mensaje = Mensaje & UsuarioActivo & Chr(13)
  Mensaje = Mensaje & EmpresaActiva & Chr(13)
  EnviarCorreo Val(LblGuia.Caption), "Reporte Novedad", Mensaje, email, oMail
End Sub

Private Sub CmdGuardar_Click()
  If TxtIdNovedad.Text <> "" Then
    AbrirRecorset rstUniversal, "INSERT INTO Novedades (Guia, IdNovedad, Comentarios, UsuIng, FhIngreso, FhNovedad, Solucion, UsuSol, FhSolucion, Solucionada) VALUES (" & Val(LblGuia) & ", " & Val(TxtIdNovedad.Text) & ", '" & TxtComentarios.Text & "', " & CodUsuarioActivo & ", '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', '" & Format(DPicFecha.value, "yyyy/mm/dd") & " " & Format(DPicHora.value, "h:m:s") & "','', " & CodUsuarioActivo & ",'" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "',0)", CnnPrincipal, adOpenDynamic, adLockOptimistic
    TxtIdNovedad.Text = ""
    TxtNovedad.Text = ""
    TxtComentarios.Text = ""
    VerNovedades
    Me.Height = 5460
    CmdAgregar.Enabled = True
    CmdSolucionar.Enabled = True
  Else
    MsgBox "Debe elegir una novedad para agragar", vbCritical, "Elija una novedad": TxtIdNovedad.SetFocus
  End If
End Sub

Private Sub CmdImpirmir_Click()
  If LstNovedades.ListItems.Count > 0 Then
    Mostrar_Reporte CnnPrincipal, 11, "Select*from sql_im_impnovedadsolucionada where ID=" & LstNovedades.SelectedItem, "", 2
  End If
End Sub

Private Sub CmdRegistro_Click()
  FufuLo = Val(LblGuia.Caption)
  FrmVerRegistroEnvioEmail.Show 1
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub CmdSolucionar_Click()
  'CmdSolucionar.Enabled = False
  'CmdAgregar.Enabled = False
  'FraSolucion.Visible = True
  'Me.Height = 6100
  'CmdGuardar.Caption = "Solucionar novedad"
  'TxtSolucion.SetFocus
  If LstNovedades.ListItems.Count > 0 Then
    FufuLo = LstNovedades.SelectedItem
    FrmSolucionarNovedad.Show 1
    VerNovedades
  End If
End Sub


Private Sub DPicFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub DPicHora_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub Form_Load()
FufuDo = 1
LblGuia = FufuLo
DPicHora = Time
DPicFecha = Date
If FufuDo = 1 Then
  CmdAgregar.Enabled = True
  CmdSolucionar.Enabled = True
Else
  LblMensaje.Caption = "Solo puede ver novedades"
End If
AbrirRecorset rstUniversal, "Select Guia, IdDespacho from guias where Guia=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
If rstUniversal.RecordCount > 0 Then
  LblIdDespacho.Caption = rstUniversal.Fields("IdDespacho") & ""
End If
CerrarRecorset rstUniversal

AbrirRecorset rstUniversal, "select ID, Orden from monitoreovehiculos where Orden=" & Val(LblIdDespacho.Caption), CnnPrincipal, adOpenDynamic, adLockOptimistic
If rstUniversal.RecordCount > 0 Then
  LblIdMonitoreo.Caption = rstUniversal.Fields("ID")
End If
CerrarRecorset rstUniversal
  
VerNovedades
End Sub

Sub VerNovedades()
  LstNovedades.ListItems.Clear
  AbrirRecorset rstUniversal, "SELECT Novedades.*, CausalesNovedad.NmNovedad FROM Novedades INNER JOIN CausalesNovedad ON Novedades.IdNovedad = CausalesNovedad.IdNovedad Where Guia=" & Val(LblGuia.Caption), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    If Val(rstUniversal!Solucionada) = 0 Then
      Set Item = LstNovedades.ListItems.Add(, , rstUniversal!Id, "Pendiente", "Pendiente")
    Else
      Set Item = LstNovedades.ListItems.Add(, , rstUniversal!Id, "Ok", "Ok")
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
  CerrarRecorset rstUniversal
  
  If LstNovedades.ListItems.Count > 0 Then
    AbrirRecorset rstUniversal, "Update guias set EnNovedad=1 where Guia=" & Val(LblGuia.Caption), CnnPrincipal, adOpenDynamic, adLockOptimistic
  Else
    AbrirRecorset rstUniversal, "Update guias set EnNovedad=0 where Guia=" & Val(LblGuia.Caption), CnnPrincipal, adOpenDynamic, adLockOptimistic
  End If
  VerNovedadesMonitoreo
End Sub
Sub VerNovedadesMonitoreo()
If Val(LblIdMonitoreo.Caption) <> 0 Then
  LstNovedadesMonitoreo.ListItems.Clear
  rstUniversal.Open "SELECT NovedadesMonitoreo.*, CausalesNovedadMonitoreo.NmNovedad FROM NovedadesMonitoreo INNER JOIN CausalesNovedadMonitoreo ON NovedadesMonitoreo.IdNovedad = CausalesNovedadMonitoreo.IdNovedad Where IdMonitoreo=" & Val(LblIdMonitoreo.Caption), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    If Val(rstUniversal.Fields("Solucionada")) = 0 Then
      Set Item = LstNovedadesMonitoreo.ListItems.Add(, , rstUniversal!Id, "Pendiente", "Pendiente")
    Else
      Set Item = LstNovedadesMonitoreo.ListItems.Add(, , rstUniversal!Id, "Ok", "Ok")
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
End If
End Sub
Private Sub TxtIdNovedad_GotFocus()
  EnfocarT TxtIdNovedad
End Sub

Private Sub TxtIdNovedad_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirConsultaGral "IdNovedad", "NmNovedad", "CausalesNovedad", CnnPrincipal
    TxtIdNovedad.Text = Principal.ToolConsultas1.DatLo
  End If
End Sub

Private Sub TxtIdNovedad_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdNovedad_Validate(Cancel As Boolean)
  AbrirRecorset rstUniversal, "select IdNovedad, NmNovedad from CausalesNovedad where IdNovedad=" & Val(TxtIdNovedad), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = False Then
    TxtNovedad = rstUniversal!NmNovedad
  Else
    TxtIdNovedad.Text = "": TxtNovedad.Text = ""
  End If
  CerrarRecorset rstUniversal
End Sub
