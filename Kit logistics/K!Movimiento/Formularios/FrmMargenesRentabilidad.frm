VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMargenesRentabilidad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Margenes de rentabilidad"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDatos 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   8415
      Begin VB.ComboBox CboCapacidad 
         Height          =   315
         ItemData        =   "FrmMargenesRentabilidad.frx":0000
         Left            =   2640
         List            =   "FrmMargenesRentabilidad.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtNmRuta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   600
         Width           =   6375
      End
      Begin VB.TextBox TxtIdRuta 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1080
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TxtCampo 
         BackColor       =   &H00FFFFFF&
         Height          =   1140
         Index           =   3
         Left            =   1080
         MaxLength       =   15
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   960
         Width           =   7215
      End
      Begin VB.TextBox TxtCampo 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   6600
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TxtCampo 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   5160
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtCampo 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Margen:"
         Height          =   195
         Left            =   4440
         TabIndex        =   14
         Top             =   240
         Width           =   585
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Index           =   7
         Left            =   600
         TabIndex        =   12
         Top             =   600
         Width           =   390
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Comentarios:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   915
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Flete:"
         Height          =   195
         Index           =   5
         Left            =   6120
         TabIndex        =   10
         Top             =   240
         Width           =   390
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Peso:"
         Height          =   195
         Index           =   2
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   405
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   210
      End
   End
   Begin MSComctlLib.Toolbar ToolMargenRentabilidad 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8640
      _ExtentX        =   15240
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
Attribute VB_Name = "FrmMargenesRentabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstMergenRentabilidad As New ADODB.Recordset
Dim Editando As Boolean
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  LaTecla (KeyCode), ToolMargenRentabilidad
End Sub
Private Sub Form_Load()
  IconosTool ToolMargenRentabilidad, Principal.IgListTool
  rstMergenRentabilidad.CursorLocation = adUseServer
  AbrirRecorset rstMergenRentabilidad, "SELECT*From Destinatarios", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Asignar rstMergenRentabilidad
End Sub
Private Sub Asignar(rstAsignar As ADODB.Recordset)
  For II = 0 To 4
    TxtCampo(II) = rstAsignar.Fields(II) & ""
  Next
End Sub
Private Sub limpiar()
  For II = 0 To 4
    TxtCampo(II).Text = ""
  Next
End Sub
Private Sub Desbloquear()
  FraDatos.Enabled = True
  BotTool 3, 17, ToolMargenRentabilidad, True
End Sub
Private Sub Bloquear()
  FraDatos.Enabled = False
  BotTool 3, 17, ToolMargenRentabilidad, False
End Sub
Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
      If Principal.ToolConsultas1.AbrirDevDatos("Digite ID", "Digite la identificacion del destinatario", 2, 0) = True Then
        FufuSt = Principal.ToolConsultas1.DatSt
        If ExRecorset("Select IdDestinatario from Destinatarios where IdDestinatario='" & FufuSt & "'") = False Then
          Desbloquear
          limpiar
          TxtCampo(0).Text = FufuSt
        Else
          MsgBox "Ya hay un destinatario creado con esta identificacion, no se pueden crear dos datos con esta identificacion", vbCritical, "El destinatario ya existe"
        End If
      End If
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            Bloquear
            AbrirRecorset rstUniversal, "Update Destinatarios set NmDestinatario='" & TxtCampo(1).Text & "', DirDestinatario='" & TxtCampo(2).Text & "', TelDestinatario='" & TxtCampo(3).Text & "', IdCiuDestinatario=" & TxtCampo(4).Text & " where IdDestinatario='" & TxtCampo(0).Text & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          AbrirRecorset rstUniversal, "INSERT INTO Destinatarios VALUES ('" & TxtCampo(0).Text & "', '" & TxtCampo(1).Text & "', '" & TxtCampo(2).Text & "', '" & TxtCampo(3).Text & "', " & Val(TxtCampo(4).Text) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
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
        Asignar rstMergenRentabilidad
        Bloquear
      End If
    Case 9  'Buscar
    
      If Principal.ToolConsultas1.AbrirDevConsulta(9, CnnPrincipal) = True Then
        AbrirRecorset rstUniversal, "Select*from Destinatarios where IdDestinatario='" & Principal.ToolConsultas1.DatSt & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          Asignar rstUniversal
        Else
          MsgBox "No se enconto el destinatario, puede ser un error interno del sistema", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If
      
    Case 11 'Primero
      UPrimero rstMergenRentabilidad
      Asignar rstMergenRentabilidad
    Case 12 'Anterior
      UAnterior rstMergenRentabilidad
      Asignar rstMergenRentabilidad
    Case 13 'Siguiente
      USiguiente rstMergenRentabilidad
      Asignar rstMergenRentabilidad
    Case 14 'Ultimo
      UUltimo rstMergenRentabilidad
      Asignar rstMergenRentabilidad
    Case 16 'Cerrar
      CerrarRecorset rstMergenRentabilidad
      FufuSt = TxtCampo(0)
      Unload Me
    Case 17 'Actualizar
      rstMergenRentabilidad.Requery
    Case 18 'Imprimir
    Case 19
      'If Val(TxtCampo(4).Text) <> 0 Then TxtNmCiudad.Text = DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & TxtCampo(4), "NmCiudad", CnnPrincipal)
  End Select
End Sub
Function Validacion() As Boolean
  If TxtCampo(1).Text <> "" Then
    If TxtCampo(2).Text <> "" Then
      If Val(TxtCampo(4).Text) <> 0 Then
        Validacion = True
      Else
        Validacion = False: MsgTit "El dato basico debe tener una ciudad": TxtCampo(4).SetFocus
      End If
    Else
      Validacion = False: MsgTit "El destinatario debe tener una direccion": TxtCampo(2).SetFocus
    End If
  Else
    Validacion = False: MsgTit "El destinatario debe tener un nombre": TxtCampo(1).SetFocus
  End If
End Function
Private Sub ToolMargenRentabilidad_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub

Private Sub TxtCampo_GotFocus(Index As Integer)
  EnfocarT TxtCampo(Index)
End Sub

Private Sub TxtCampo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    If Index = 4 Then
      Principal.ToolConsultas1.AbrirDevConsulta 1, CnnPrincipal
      TxtCampo(4).Text = Principal.ToolConsultas1.DatLo
    End If
  End If
End Sub

Private Sub TxtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
    Case 3
      ValidarEntrada TxtCampo(1), KeyAscii, 1
    Case 4
      ValidarEntrada TxtCampo(1), KeyAscii, 1
  End Select
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtCampo_Validate(Index As Integer, Cancel As Boolean)
  If Index = 4 Then
    If Val(TxtCampo(4).Text) <> 0 Then
      AbrirRecorset rstUniversal, "SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Val(TxtCampo(4).Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        'TxtNmCiudad = rstUniversal!NmCiudad & ""
      Else
        'TxtNmCiudad = "": TxtCampo(4) = ""
      End If
    End If
  End If
End Sub


