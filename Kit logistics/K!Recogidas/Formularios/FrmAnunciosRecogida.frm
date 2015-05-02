VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAnunciosRecogida 
   Caption         =   "Anuncios - Recogidas"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7290
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdCorregirUnidades 
      Caption         =   "Corregir unidades"
      Height          =   255
      Left            =   960
      TabIndex        =   49
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton CmdInfoAsignacion 
      Caption         =   "Info asignacion"
      Height          =   255
      Left            =   3960
      TabIndex        =   44
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Frame FraEfectiva 
      Enabled         =   0   'False
      Height          =   550
      Left            =   960
      TabIndex        =   33
      Top             =   6000
      Width           =   7455
      Begin VB.TextBox TxtHrefectiva 
         Height          =   285
         Left            =   6360
         TabIndex        =   43
         Top             =   170
         Width           =   975
      End
      Begin VB.TextBox TxtFhEfectiva 
         Height          =   285
         Left            =   3720
         TabIndex        =   40
         Top             =   170
         Width           =   1455
      End
      Begin VB.CheckBox ChkProgramada 
         Alignment       =   1  'Right Justify
         Caption         =   "Programada"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   170
         Width           =   1215
      End
      Begin VB.CheckBox ChkEfectiva 
         Alignment       =   1  'Right Justify
         Caption         =   "Efectiva"
         Height          =   255
         Left            =   1560
         TabIndex        =   34
         Top             =   170
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hr Efectiva:"
         Height          =   195
         Left            =   5400
         TabIndex        =   42
         Top             =   165
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fh Efectiva:"
         Height          =   195
         Left            =   2760
         TabIndex        =   41
         Top             =   170
         Width           =   855
      End
   End
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   3135
      Left            =   960
      TabIndex        =   27
      Top             =   2880
      Width           =   7455
      Begin VB.TextBox TxtMotivoCancelacion 
         Enabled         =   0   'False
         Height          =   615
         Left            =   960
         TabIndex        =   50
         Top             =   2400
         Width           =   6375
      End
      Begin VB.TextBox TxtIdVehiculo 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TxtIdConductor 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtComentarios 
         Height          =   1005
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Top             =   1320
         Width           =   5055
      End
      Begin VB.TextBox TxtRuta 
         Height          =   285
         Left            =   960
         MaxLength       =   5
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DPHora 
         Height          =   285
         Left            =   5880
         TabIndex        =   8
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "hh:mm:ss"
         Format          =   16908290
         UpDown          =   -1  'True
         CurrentDate     =   38510
      End
      Begin VB.TextBox TxtUnidades 
         Alignment       =   1  'Right Justify
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
         Left            =   960
         MaxLength       =   5
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox TxtKilosReales 
         Alignment       =   1  'Right Justify
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
         Left            =   960
         MaxLength       =   5
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TxtKilosVol 
         Alignment       =   1  'Right Justify
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
         Left            =   960
         MaxLength       =   5
         TabIndex        =   11
         Top             =   2040
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DPFHRecogida 
         Height          =   285
         Left            =   3360
         TabIndex        =   7
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   16908291
         CurrentDate     =   38244
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Conductor:"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Vehiculo:"
         Height          =   195
         Left            =   255
         TabIndex        =   47
         Top             =   960
         Width           =   660
      End
      Begin VB.Label LblConsulta 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   46
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label LblConsulta 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   37
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Left            =   525
         TabIndex        =   36
         Top             =   600
         Width           =   390
      End
      Begin VB.Label LblEtiquetas 
         AutoSize        =   -1  'True
         Caption         =   "Fh Recogida:"
         Height          =   195
         Index           =   0
         Left            =   2280
         TabIndex        =   32
         Top             =   960
         Width           =   960
      End
      Begin VB.Label LblEtiquetas 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         Height          =   195
         Index           =   1
         Left            =   5400
         TabIndex        =   31
         Top             =   960
         Width           =   390
      End
      Begin VB.Label LblEtiquetas 
         AutoSize        =   -1  'True
         Caption         =   "Unidades:"
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   30
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label LblEtiquetas 
         AutoSize        =   -1  'True
         Caption         =   "Peso Real:"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   29
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label LblEtiquetas 
         AutoSize        =   -1  'True
         Caption         =   "Peso Vol:"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   28
         Top             =   2040
         Width           =   675
      End
   End
   Begin VB.Frame FraVeCon 
      Enabled         =   0   'False
      Height          =   550
      Left            =   6240
      TabIndex        =   25
      Top             =   6600
      Width           =   2175
      Begin VB.TextBox TxtAsignacion 
         Height          =   285
         Left            =   960
         MaxLength       =   6
         TabIndex        =   12
         Top             =   160
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Asigancion:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   165
         Width           =   825
      End
   End
   Begin VB.Frame FraCliente 
      Enabled         =   0   'False
      Height          =   1335
      Left            =   960
      TabIndex        =   21
      Top             =   1440
      Width           =   7455
      Begin VB.TextBox TxtTelAnunciante 
         Height          =   285
         Left            =   6120
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TxtDirAnunciante 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox TxtAnunciante 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   6255
      End
      Begin VB.TextBox TxtIdCliente 
         Height          =   285
         Left            =   1080
         MaxLength       =   11
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LblEtiquetas 
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   39
         Top             =   960
         Width           =   720
      End
      Begin VB.Label LblEtiquetas 
         AutoSize        =   -1  'True
         Caption         =   "Tel:"
         Height          =   195
         Index           =   7
         Left            =   5760
         TabIndex        =   38
         Top             =   960
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   450
         TabIndex        =   24
         Top             =   240
         Width           =   525
      End
      Begin VB.Label LblConsulta 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   23
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label LblEtiquetas 
         AutoSize        =   -1  'True
         Caption         =   "Anunciante:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame FraDatosLectura 
      Enabled         =   0   'False
      Height          =   615
      Left            =   960
      TabIndex        =   14
      Top             =   720
      Width           =   7455
      Begin VB.TextBox TxtEstado 
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
         Left            =   5520
         TabIndex        =   17
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TxtAnuncio 
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtFecha 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2760
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Anuncio:"
         Height          =   195
         Index           =   12
         Left            =   360
         TabIndex        =   20
         Top             =   240
         Width           =   630
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   19
         Left            =   4920
         TabIndex        =   19
         Top             =   240
         Width           =   540
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Index           =   8
         Left            =   2280
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
   End
   Begin MSComctlLib.Toolbar ToolAnuncios 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1005
      ButtonWidth     =   847
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
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
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Car"
            Object.ToolTipText     =   "Carga informacion adicional [Pausa]"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmAnunciosRecogida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Editando As Boolean
Dim rstAnuncios As New ADODB.Recordset

Private Sub CmdCorregirUnidades_Click()
  Dim UnidadesN As Double
End Sub

Private Sub CmdInfoAsignacion_Click()
  If Val(TxtAsignacion) <> 0 Then
    FufuLo = Val(TxtAsignacion.Text)
    FrmInfoAsignacion.Show 1
  End If
End Sub
Private Sub DPFHRecogida_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub
Private Sub DPHora_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  LaTecla (KeyCode), ToolAnuncios
End Sub
Private Sub Form_Load()
  IconosTool ToolAnuncios, Principal.IgListTool
  rstAnuncios.CursorLocation = adUseServer
  AbrirRecorset rstAnuncios, "SELECT*From Anuncios", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Asignar
End Sub
Sub Desbloquear()
  BotTool 3, 17, ToolAnuncios, True
  FraCliente.Enabled = True
  FraDatos.Enabled = True
  CmdInfoAsignacion.Enabled = False
  TxtComentarios.Locked = False
End Sub
Sub Bloquear()
  BotTool 3, 17, ToolAnuncios, False
  FraCliente.Enabled = False
  FraDatos.Enabled = False
  CmdInfoAsignacion.Enabled = True
  TxtComentarios.Locked = True
End Sub
Sub Asignar()
  TxtAnuncio.Text = rstAnuncios!IdAnuncio
  TxtFecha.Text = rstAnuncios!FhAnuncio
  TxtEstado.Text = DevEstadoDespacho(rstAnuncios!Estado)
  TxtIdCliente.Text = rstAnuncios!IdCliente & ""
  TxtAnunciante.Text = rstAnuncios!Anunciante & ""
  TxtDirAnunciante.Text = rstAnuncios!DirAnunciante & ""
  TxtTelAnunciante.Text = rstAnuncios!TelAnunciante & ""
  TxtAsignacion.Text = rstAnuncios!Asignacion & ""
  TxtRuta.Text = rstAnuncios!IdRuta & ""
  DPHora.Value = Format(rstAnuncios!FhRecogida, "h:m:s")
  DPFHRecogida.Value = Format(rstAnuncios!FhRecogida, "dd/mm/yy")
  TxtUnidades.Text = rstAnuncios!Unidades
  TxtKilosReales.Text = rstAnuncios!KilosReales
  TxtKilosVol.Text = rstAnuncios!KilosVol
  TxtComentarios.Text = rstAnuncios!Comentarios & ""
  TxtMotivoCancelacion.Text = rstAnuncios!MotivoCancelacion & ""
  ChkEfectiva.Value = DevCheck(rstAnuncios!Efectiva)
  ChkProgramada.Value = DevCheck(rstAnuncios!Programada)
  TxtFhEfectiva.Text = Format(rstAnuncios!TiempoEfectiva & "", "dd/mm/yyyy")
  TxtHrefectiva.Text = Format(rstAnuncios!TiempoEfectiva & "", "hh:mm:ss")
  TxtIdConductor.Text = rstAnuncios!IdConductor & ""
  TxtIdVehiculo.Text = rstAnuncios!IdVehiculo & ""
  LimpiarConsulta
End Sub
Sub limpiar()
  TxtAnuncio.Text = ""
  TxtFecha.Text = ""
  TxtEstado.Text = ""
  TxtIdCliente.Text = ""
  TxtAnunciante.Text = ""
  TxtDirAnunciante.Text = ""
  TxtTelAnunciante.Text = ""
  TxtAsignacion.Text = ""
  TxtRuta.Text = ""
  DPHora.Value = Time
  DPFHRecogida.Value = Date
  TxtUnidades.Text = ""
  TxtKilosReales.Text = ""
  TxtKilosVol.Text = ""
  TxtComentarios.Text = ""
  TxtFhEfectiva.Text = ""
  TxtHrefectiva.Text = ""
  ChkEfectiva.Value = 0
  ChkProgramada.Value = 0
  TxtIdConductor.Text = ""
  TxtIdVehiculo.Text = ""
  TxtMotivoCancelacion.Text = ""
  LimpiarConsulta
End Sub
Sub LimpiarConsulta()
  LblConsulta(1).Caption = ""
  LblConsulta(2).Caption = ""
End Sub
Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
      Desbloquear
      limpiar
      TxtIdCliente.SetFocus
      TxtFecha.Text = Date
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            Bloquear
            AbrirRecorset rstUniversal, "Update Anuncios set IdCliente='" & TxtIdCliente.Text & "', Anunciante='" & TxtAnunciante & "', DirAnunciante='" & TxtDirAnunciante.Text & "', TelAnunciante='" & TxtTelAnunciante.Text & "', IdRuta=" & Val(TxtRuta) & ", Vehiculo='" & TxtAsignacion.Text & "', HoraRec='" & DPHora.Value & "', FhRecogida='" & DPFHRecogida.Value & "', Unidades=" & Val(TxtUnidades) & ", KilosReales=" & Val(TxtKilosReales) & ", KilosVol=" & Val(TxtKilosVol.Text) & ", Comentarios='" & TxtComentarios & "' where IdAnuncio=" & Val(TxtAnuncio), CnnPrincipal, adOpenDynamic, adLockOptimistic
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          TxtAnuncio.Text = SacarConsecutivo("Anuncios")
          AbrirRecorset rstUniversal, "INSERT INTO Anuncios (IdAnuncio, FhAnuncio, IdCliente, Anunciante, DirAnunciante, TelAnunciante, IdRuta, FhRecogida, Unidades, KilosReales, KilosVol, Comentarios, Programada, Estado, Efectiva, Coperaciones, Orden, IdVehiculo, IdConductor, IdEmpresa) " & _
          " VALUES (" & Val(TxtAnuncio) & ",'" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "','" & TxtIdCliente & "','" & TxtAnunciante & "','" & TxtDirAnunciante.Text & "','" & TxtTelAnunciante.Text & "'," & Val(TxtRuta.Text) & ",'" & Format(DPFHRecogida.Value, "yyyy/mm/dd") & " " & Format(DPHora.Value, "h:m:s") & "'," & Val(TxtUnidades) & "," & Val(TxtKilosReales) & "," & Val(TxtKilosVol) & ",'" & TxtComentarios & "',0,'P',0," & Coperaciones & ",0,'" & TxtIdVehiculo.Text & "','" & TxtIdConductor.Text & "',1)", CnnPrincipal, adOpenDynamic, adLockOptimistic

          TxtEstado = DevEstadoDespacho("D")
          Bloquear
          CmdInfoAsignacion.SetFocus
        End If
      End If
    Case 5  'Editar
      Editando = True
      Desbloquear
      LimpiarConsulta
      TxtIdCliente.SetFocus
    Case 6 'Eliminar
      If rstAnuncios!Estado & "" = "D" Then
        If MsgBox("¿Esta seguro de que desea ELIMINAR el anuncio?", vbQuestion + vbYesNo, "Eliminar anuncio") = vbYes Then
          AbrirRecorset rstUniversal, "Delete Anuncios where IdAnuncio=" & Val(TxtAnuncio), CnnPrincipal, adOpenDynamic, adLockOptimistic
        End If
        AccionTool 11
      End If
    Case 7 'Cancelar
      If Editando = True Then
        AccionTool 4
      Else
        Asignar
        Bloquear
      End If
    Case 9  'Buscar
      'If Principal.ToolConsultas1.AbrirDevDatos("Numero de anuncio", "Digite del anuncio que desea buscar", 3) = True Then If BuscaRegistro("IdAnuncio=" & Principal.ToolConsultas1.DatLo, rstAnuncios) = True Then Asignar
    Case 11 'Primero
      UPrimero rstAnuncios
      Asignar
    Case 12 'Anterior
      UAnterior rstAnuncios
      Asignar
    Case 13 'Siguiente
      USiguiente rstAnuncios
      Asignar
    Case 14 'Ultimo
      UUltimo rstAnuncios
      Asignar
    Case 16 'Cerrar
      CerrarRecorset rstAnuncios
      Unload Me
    Case 17 'Actualizar
      rstAnuncios.Requery
    Case 18 'Imprimir
     Mostrar_Reporte CnnPrincipal, 17, "Select*from sql_im_impanuncio where IdAnuncio=" & Val(TxtAnuncio.Text), "", 2
     
    Case 19 'Cargar
      If TxtIdCliente.Text <> "" Then LblConsulta(1) = DevNombreDatosBasicos(TxtIdCliente.Text)
      If Val(TxtRuta) <> 0 Then LblConsulta(2) = DevResBus("SELECT IdRutaRec, NmRuta From RutasUrbanas Where IdRutaRec =" & Val(TxtRuta), "NmRuta")
  End Select
End Sub

Function Validacion() As Boolean
  If TxtIdCliente.Text <> "" Then
    If TxtAnunciante.Text <> "" Then
      If TxtDirAnunciante.Text <> "" Then
        If TxtTelAnunciante.Text <> "" Then
          If DPFHRecogida.Value >= Date Then
              Validacion = True
          Else
            Validacion = False: MsgTit "No se puede programar una recogida o hacer un anuncio para un dia anterior a hoy": DPFHRecogida.SetFocus
          End If
        Else
          Validacion = False: MsgTit "El anuncio debe tener un telefono para contactar el anunciante": TxtTelAnunciante.SetFocus
        End If
      Else
        Validacion = False: MsgTit "El anuncio debe tener una direccion para recoger la mercancía": TxtDirAnunciante.SetFocus
      End If
    Else
      Validacion = False: MsgTit "El anuncio debe tener un anunciante": TxtAnunciante.SetFocus
    End If
  Else
    Validacion = False: MsgTit "El anuncio debe tener un cliente, en el caso de que no lo tenga coloque 0": TxtIdCliente.SetFocus
  End If
End Function
Private Sub ToolAnuncios_ButtonClick(ByVal Button As MSComctlLib.Button)
  AccionTool Button.Index
End Sub
Private Sub TxtAnunciante_GotFocus()
  EnfocarT TxtAnunciante
End Sub
Private Sub TxtAnunciante_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub TxtComentarios_GotFocus()
 EnfocarT TxtComentarios
End Sub
Private Sub TxtComentarios_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then KeyAscii = 0
End Sub
Private Sub TxtDirAnunciante_GotFocus()
  EnfocarT TxtDirAnunciante
End Sub
Private Sub TxtDirAnunciante_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub



Private Sub TxtIdCliente_GotFocus()
  EnfocarT TxtIdCliente
End Sub
Private Sub TxtIdCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsulta 7, CnnPrincipal
    TxtIdCliente.Text = Principal.ToolConsultas1.DatSt
  End If
End Sub
Private Sub TxtIdCliente_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtIdCliente, KeyAscii, 1
End Sub
Private Sub TxtIdCliente_Validate(Cancel As Boolean)
  If TxtIdCliente.Text = "" Then TxtIdCliente.Text = "0"
    If Val(TxtIdCliente) <> 0 Then
      'AbrirRecorset rstUniversal, "SELECT Idcliente, NmCliente, DirCliente, TelCliente, Recoge, HorarioRecoge, DirBodega, IdRutaRecogida, EncargadoRec From Clientes Where IdCliente ='" & TxtIdCliente & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      AbrirRecorset rstUniversal, "SELECT IdTercero, RazonSocial, Direccion, Telefono From Terceros Where IdTercero ='" & TxtIdCliente & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        LblConsulta(1) = rstUniversal!RazonSocial & ""
        TxtDirAnunciante = rstUniversal!Direccion & ""
        TxtTelAnunciante = rstUniversal!Telefono & ""
        TxtAnunciante.Text = rstUniversal!RazonSocial & ""
        'TxtRuta = rstUniversal!IdRutaRecogida & ""
        'DPHora.Value = rstUniversal!HorarioRecoge
        'If rstUniversal!Recoge = 0 Then MsgBox "Cuidado, este cliente no tiene habilitada la casilla de que se le recoge", vbExclamation, "No esta marcado que se le recoge"
        'If Val(TxtRuta) <> 0 Then LblConsulta(2) = DevResBus("SELECT IdRutaRec, NmRuta From RutasUrbanas Where IdRutaRec =" & Val(TxtRuta), "NmRuta")
      Else
        LblConsulta(1).Caption = "": TxtIdCliente.Text = ""
      End If
      CerrarRecorset rstUniversal
    End If
End Sub

Private Sub TxtIdVehiculo1_Change()

End Sub

Private Sub TxtIdConductor_GotFocus()
  EnfocarT TxtIdConductor
End Sub

Private Sub TxtIdConductor_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsulta 3, CnnPrincipal
    TxtIdConductor.Text = Principal.ToolConsultas1.DatSt
  End If
End Sub

Private Sub TxtIdConductor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtIdConductor, KeyAscii, 1
End Sub

Private Sub TxtIdConductor_Validate(Cancel As Boolean)
  If TxtIdConductor.Text <> "" Then
    AbrirRecorset rstUniversal, "Select IdConductor, Concat(Nombre, ' ', Apellido1,  ' ', Apellido2) as NmConductor, FhVenceLic, ConductorInactivo From Conductores where IdConductor='" & TxtIdConductor.Text & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      If CDate(rstUniversal.Fields("FhVenceLic")) < Date Then
        MsgTit "El conductor no puede viajar porque tiene la licencia vencida o esta inactivo"
        TxtIdConductor.Text = "": LblConsulta(0).Caption = ""
      Else
        If Val(rstUniversal.Fields("ConductorInactivo")) = 0 Then
          LblConsulta(0).Caption = rstUniversal.Fields("NmConductor")
        Else
          MsgTit "El conductor no puede viajar porque esta inactivo"
          TxtIdConductor.Text = "": LblConsulta(0).Caption = ""
        End If
      End If
    Else
      MsgBox "El conductor no existe", vbCritical
      LblConsulta(0).Caption = "": TxtIdConductor.Text = ""
    End If
    CerrarRecorset rstUniversal
  End If
End Sub

Private Sub TxtIdVehiculo_GotFocus()
  EnfocarT TxtIdVehiculo
End Sub

Private Sub TxtIdVehiculo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsulta 5, CnnPrincipal
    TxtIdVehiculo.Text = Principal.ToolConsultas1.DatSt
  End If
End Sub

Private Sub TxtIdVehiculo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdVehiculo_Validate(Cancel As Boolean)
      AbrirRecorset rstUniversal, "Select IdPlaca, VenceSoat, Inactivo from Vehiculos where IdPlaca='" & TxtIdVehiculo.Text & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.EOF = True Then
        MsgTit "El vehiculo no existe"
        TxtIdVehiculo.Text = ""
      Else
        If CDate(rstUniversal.Fields("VenceSoat")) < Date Or Val(rstUniversal.Fields("Inactivo")) = 1 Then
          MsgTit "El vehiculo esta inactivo o el soat estan vencidos, no se puede despachar este vehiculo"
          TxtIdVehiculo.Text = ""
        End If
      End If
      CerrarRecorset rstUniversal
End Sub

Private Sub TxtKilosReales_GotFocus()
  EnfocarT TxtKilosReales
End Sub
Private Sub TxtKilosReales_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtKilosReales, KeyAscii, 1
End Sub
Private Sub TxtKilosVol_GotFocus()
  EnfocarT TxtKilosVol
End Sub
Private Sub TxtKilosVol_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtKilosVol, KeyAscii, 1
End Sub
Private Sub TxtRuta_GotFocus()
  EnfocarT TxtRuta
End Sub
Private Sub TxtRuta_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsultaCO 1, Coperaciones, CnnPrincipal
    TxtRuta.Text = Principal.ToolConsultas1.DatLo
  End If
End Sub
Private Sub TxtRuta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtRuta, KeyAscii, 1
End Sub
Private Sub TxtRuta_Validate(Cancel As Boolean)
  If Val(TxtRuta) <> 0 Then
    AbrirRecorset rstUniversal, "SELECT IdRutaRec, NmRuta FROM RutasUrbanas where IdRutaRec=" & Val(TxtRuta), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      LblConsulta(2) = rstUniversal!NmRuta & ""
    Else
      LblConsulta(2) = "": TxtRuta = ""
    End If
    CerrarRecorset rstUniversal
  End If
End Sub
Private Sub TxtTelAnunciante_GotFocus()
  EnfocarT TxtTelAnunciante
End Sub
Private Sub TxtTelAnunciante_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtTelAnunciante, KeyAscii, 1
End Sub
Private Sub TxtUnidades_GotFocus()
  EnfocarT TxtUnidades
End Sub
Private Sub TxtUnidades_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtUnidades, KeyAscii, 1
End Sub
