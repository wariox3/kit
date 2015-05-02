VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRemitentes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Remitentes..."
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8370
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   8055
      Begin VB.TextBox TxtTelRemitente 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox TxtIdRemitente 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TxtRemitente 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   0
         Top             =   600
         Width           =   6615
      End
      Begin VB.TextBox TxtDirRemitente 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   1
         Top             =   960
         Width           =   6615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Telefono:"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
         Height          =   195
         Left            =   480
         TabIndex        =   7
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Remitente:"
         Height          =   195
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   540
      End
   End
   Begin MSComctlLib.Toolbar ToolRemitentes 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8370
      _ExtentX        =   14764
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
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Acc"
            Style           =   5
            Object.Width           =   1e-4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmRemitentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstRemitentes As New ADODB.Recordset
Dim Editando As Boolean
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  LaTecla (KeyCode), ToolRemitentes
End Sub
Private Sub Form_Load()
  IconosTool ToolRemitentes, Principal.IgListTool
  rstRemitentes.CursorLocation = adUseClient
  AbrirRecorset rstRemitentes, "SELECT*From Remitentes", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Asignar rstRemitentes
End Sub
Private Sub Asignar(rstAsignar As ADODB.Recordset)
  TxtIdRemitente.Text = rstAsignar!IdRemitente
  TxtRemitente.Text = rstAsignar!NmRemitente
  TxtDirRemitente.Text = rstAsignar!DirRemitente
  TxtTelRemitente.Text = rstAsignar!TelRemitente
End Sub
Private Sub limpiar()
  TxtIdRemitente.Text = ""
  TxtRemitente.Text = ""
  TxtDirRemitente.Text = ""
  TxtTelRemitente.Text = ""
End Sub
Private Sub Desbloquear()
  FraDatos.Enabled = True
  BotTool 3, 17, ToolRemitentes, True
End Sub
Private Sub Bloquear()
  FraDatos.Enabled = False
  BotTool 3, 17, ToolRemitentes, False
End Sub
Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
      If Principal.ToolConsultas1.AbrirDevDatos("Digite ID", "Digite la identificacion del remitente", 2, 0) = True Then
        FufuSt = Principal.ToolConsultas1.DatSt
        If ExRecorset("Select IdRemitente from remitentes where IdRemitente='" & FufuSt & "'") = False Then
          Desbloquear
          limpiar
          TxtIdRemitente = FufuSt
        Else
          MsgBox "Ya hay un remitente creado con este codigo", vbCritical, "El remitente ya existe"
        End If
      End If
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            Bloquear
            AbrirRecorset rstUniversal, "Update Remitentes set NmRemitente='" & TxtRemitente & "', DirRemitente='" & TxtDirRemitente & "', TelRemitente='" & TxtTelRemitente & "' where IdRemitente='" & TxtIdRemitente & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          TxtIdRemitente = SacarConsecutivo("Remitentes", CnnPrincipal)
          AbrirRecorset rstUniversal, "INSERT INTO Remitentes VALUES (" & TxtIdRemitente & ", '" & TxtRemitente.Text & "', '" & TxtDirRemitente.Text & "', '" & TxtTelRemitente & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
          Bloquear
        End If
      End If
    Case 5  'Editar
      Editando = True
      Desbloquear
    Case 6 'Eliminar
      
    Case 7 'Cancelar
      If Editando = True Then
        AccionTool 4
      Else
        Asignar rstRemitentes
        Bloquear
      End If
    Case 9  'Buscar
      If Principal.ToolConsultas1.AbrirDevConsulta(4, CnnPrincipal) = True Then
        AbrirRecorset rstUniversal, "Select*from Remitentes where IdRemitente='" & Principal.ToolConsultas1.DatSt & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          Asignar rstUniversal
        Else
          MsgBox "No se enconto el remiente, puede ser un error interno del sistema", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If
      
    Case 11 'Primero
      UPrimero rstRemitentes
      Asignar rstRemitentes
    Case 12 'Anterior
      UAnterior rstRemitentes
      Asignar rstRemitentes
    Case 13 'Siguiente
      USiguiente rstRemitentes
      Asignar rstRemitentes
    Case 14 'Ultimo
      UUltimo rstRemitentes
      Asignar rstRemitentes
    Case 16 'Cerrar
      CerrarRecorset rstRemitentes
      FufuSt = TxtIdRemitente.Text
      Unload Me
    Case 17 'Actualizar
      rstRemitentes.Requery
    Case 18 'Imprimir
  End Select
End Sub
Function Validacion() As Boolean
  If TxtRemitente.Text <> "" Then
    If TxtDirRemitente.Text <> "" Then
      If TxtTelRemitente.Text <> "" Then
        Validacion = True
      Else
        Validacion = False: MsgTit "El remitente debe tener un telefono": TxtTelRemitente.SetFocus
      End If
    Else
      Validacion = False: MsgTit "El remitente debe tener una direccion": TxtDirRemitente.SetFocus
    End If
  Else
    Validacion = False: MsgTit "El remitente debe tener un nombre": TxtRemitente.SetFocus
  End If
End Function
Private Sub ToolRemitentes_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub

Private Sub TxtDirRemitente_GotFocus()
  EnfocarT TxtDirRemitente
End Sub

Private Sub TxtDirRemitente_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub TxtRemitente_GotFocus()
  EnfocarT TxtRemitente
End Sub

Private Sub TxtRemitente_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub TxtTelRemitente_GotFocus()
  EnfocarT TxtTelRemitente
End Sub

Private Sub TxtTelRemitente_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtTelRemitente, KeyAscii, 1
End Sub
