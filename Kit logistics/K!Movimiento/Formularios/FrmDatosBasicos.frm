VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDatosBasicos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Datos Basicos..."
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8175
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   2895
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   7335
      Begin VB.TextBox TxtNmCiudad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Top             =   2400
         Width           =   5055
      End
      Begin VB.TextBox TxtCampo 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   7
         Left            =   1080
         TabIndex        =   16
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox TxtCampo 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   6
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   15
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox TxtCampo 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         Left            =   1080
         TabIndex        =   14
         Top             =   1680
         Width           =   5895
      End
      Begin VB.TextBox TxtCampo 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         Left            =   1080
         TabIndex        =   13
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox TxtCampo 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   1080
         TabIndex        =   12
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TxtCampo 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   1080
         TabIndex        =   11
         Top             =   600
         Width           =   5895
      End
      Begin VB.TextBox TxtCampo 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   3600
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TxtCampo 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Index           =   7
         Left            =   480
         TabIndex        =   9
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Telefono:"
         Height          =   195
         Index           =   6
         Left            =   345
         TabIndex        =   8
         Top             =   2040
         Width           =   675
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
         Height          =   195
         Index           =   5
         Left            =   300
         TabIndex        =   7
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Apellido 2:"
         Height          =   195
         Index           =   4
         Left            =   285
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Apellido 1:"
         Height          =   195
         Index           =   3
         Left            =   285
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   2
         Left            =   420
         TabIndex        =   4
         Top             =   600
         Width           =   600
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "TP ID:"
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   465
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         Height          =   195
         Index           =   0
         Left            =   810
         TabIndex        =   2
         Top             =   240
         Width           =   210
      End
   End
   Begin MSComctlLib.Toolbar ToolDatosBasicos 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
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
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar para Re-Repartir"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar para Re-Despachar"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmDatosBasicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstDatosBasicos As New ADODB.Recordset
Dim Editando As Boolean
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  LaTecla (KeyCode), ToolDatosBasicos
End Sub
Private Sub Form_Load()
  IconosTool ToolDatosBasicos, Principal.IgListTool
  rstDatosBasicos.CursorLocation = adUseServer
  AbrirRecorset rstDatosBasicos, "SELECT*From DatosBasicos", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Asignar
End Sub
Private Sub Asignar()
  For II = 0 To 7
    TxtCampo(II) = rstDatosBasicos.Fields(II) & ""
  Next
  TxtNmCiudad.Text = ""
End Sub
Private Sub limpiar()
  For II = 0 To 7
    TxtCampo(II).Text = ""
  Next
  TxtNmCiudad.Text = ""
End Sub
Private Sub Desbloquear()
  FraDatos.Enabled = True
  BotTool 3, 17, ToolDatosBasicos, True
End Sub
Private Sub Bloquear()
  FraDatos.Enabled = False
  BotTool 3, 17, ToolDatosBasicos, False
End Sub
Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
      
      If Principal.ToolConsultas1.AbrirDevDatos("Digite ID", "Digite la identificacion del dato basico", 2, 0) = True Then
        FufuSt = Principal.ToolConsultas1.DatSt
        If ExRecorset("Select ID from DatosBasicos where ID='" & FufuSt & "'") = False Then
          Desbloquear
          limpiar
          TxtCampo(0).Text = FufuSt
        Else
          MsgBox "Ya hay un Dato Basico creado con esta identificacion, no se pueden crear dos datos con esta identificacion", vbCritical, "El dato basico ya existe"
        End If
      End If
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            Bloquear
            AbrirRecorset rstUniversal, "Update DatosBasicos set TpDoc='" & TxtCampo(1).Text & "', Nombre='" & TxtCampo(2).Text & "', Apellido1='" & TxtCampo(3).Text & "', Apellido2='" & TxtCampo(4).Text & "', Direccion='" & TxtCampo(5).Text & "', Telefono='" & TxtCampo(6).Text & "', IdCiudad=" & TxtCampo(7).Text & " where ID='" & TxtCampo(0).Text & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          AbrirRecorset rstUniversal, "INSERT INTO DatosBasicos VALUES ('" & TxtCampo(0).Text & "', '" & TxtCampo(1).Text & "', '" & TxtCampo(2).Text & "', '" & TxtCampo(3).Text & "', '" & TxtCampo(4).Text & "', '" & TxtCampo(5).Text & "', '" & TxtCampo(6).Text & "', " & Val(TxtCampo(7).Text) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
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
        Asignar
        Bloquear
      End If
    Case 9  'Buscar
       'FrmBuscarDatosBasicos.Show 1
       'If BuscaRegistro("ID='" & FufuSt & "'", rstDatosBasicos) = True Then Asignar
    Case 11 'Primero
      UPrimero rstDatosBasicos
      Asignar
    Case 12 'Anterior
      UAnterior rstDatosBasicos
      Asignar
    Case 13 'Siguiente
      USiguiente rstDatosBasicos
      Asignar
    Case 14 'Ultimo
      UUltimo rstDatosBasicos
      Asignar
    Case 16 'Cerrar
      CerrarRecorset rstDatosBasicos
      FufuSt = TxtCampo(0)
      Unload Me
    Case 17 'Actualizar
      rstDatosBasicos.Requery
    Case 18 'Imprimir
    Case 19
      If Val(TxtCampo(7).Text) <> 0 Then TxtNmCiudad.Text = DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & TxtCampo(7), "NmCiudad", CnnPrincipal)
  End Select
End Sub
Function Validacion() As Boolean
  If TxtCampo(0).Text <> "" Then
    If TxtCampo(5).Text <> "" Then
      If TxtCampo(7).Text <> "" And Val(TxtCampo(7).Text) <> 0 Then
        Validacion = True
      Else
        Validacion = False: MsgTit "El dato basico debe tener una ciudad": TxtCampo(7).SetFocus
      End If
    Else
      Validacion = False: MsgTit "El dato basico debe tener una direccion": TxtCampo(5).SetFocus
    End If
  Else
    Validacion = False: MsgTit "El dato basico debe tener un nombre": TxtCampo(0).SetFocus
  End If
End Function
Private Sub ToolDatosBasicos_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub

Private Sub TxtCampo_GotFocus(Index As Integer)
  EnfocarT TxtCampo(Index)
End Sub

Private Sub TxtCampo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    If Index = 7 Then
      Principal.ToolConsultas1.AbrirDevConsulta 1, CnnPrincipal
      TxtCampo(7).Text = Principal.ToolConsultas1.DatLo
    End If
  End If
End Sub

Private Sub TxtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
    Case 1
      ValidarEntrada TxtCampo(1), KeyAscii, 4
    Case 7
      ValidarEntrada TxtCampo(1), KeyAscii, 1
  End Select
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtCampo_Validate(Index As Integer, Cancel As Boolean)
  If Index = 7 Then
    If Val(TxtCampo(7).Text) <> 0 Then
      AbrirRecorset rstUniversal, "SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & Val(TxtCampo(7).Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        TxtNmCiudad = rstUniversal!NmCiudad & ""
      Else
        TxtNmCiudad = "": TxtCampo(7) = ""
      End If
    End If
  End If
End Sub
