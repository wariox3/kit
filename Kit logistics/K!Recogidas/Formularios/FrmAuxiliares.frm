VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAuxiliares 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Auxiliares..."
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7545
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   7335
      Begin VB.TextBox TxtCO 
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TxtTelefono 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4680
         MaxLength       =   15
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TxtNmAuxiliar 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   960
         TabIndex        =   1
         Top             =   600
         Width           =   5895
      End
      Begin VB.TextBox TxtIdAuxiliar 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label LblConsulta 
         AutoSize        =   -1  'True
         Caption         =   "CO:"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   270
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Telefono:"
         Height          =   195
         Index           =   6
         Left            =   3960
         TabIndex        =   7
         Top             =   240
         Width           =   675
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   6
         Top             =   600
         Width           =   600
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         Height          =   195
         Index           =   0
         Left            =   690
         TabIndex        =   4
         Top             =   240
         Width           =   210
      End
   End
   Begin MSComctlLib.Toolbar ToolAuxiliares 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7545
      _ExtentX        =   13309
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
Attribute VB_Name = "FrmAuxiliares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstAuxiliares As New ADODB.Recordset
Dim Editando As Boolean
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  LaTecla (KeyCode), ToolAuxiliares
End Sub
Private Sub Form_Load()
  IconosTool ToolAuxiliares, Principal.IgListTool
  rstAuxiliares.CursorLocation = adUseServer
  AbrirRecorset rstAuxiliares, "SELECT*From Auxiliares", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Asignar
End Sub
Sub Asignar()
  TxtIdAuxiliar.Text = rstAuxiliares!IdAuxiliar
  TxtNmAuxiliar.Text = rstAuxiliares!NmAuxiliar & ""
  TxtTelefono.Text = rstAuxiliares!TelAuxiliar & ""
  TxtCO.Text = rstAuxiliares!COAuxiliar & ""
End Sub
Sub limpiar()
  TxtIdAuxiliar.Text = ""
  TxtNmAuxiliar.Text = ""
  TxtTelefono.Text = ""
  TxtCO.Text = ""
End Sub
Sub Desbloquear()
  FraDatos.Enabled = True
  BotTool 3, 17, ToolAuxiliares, True
End Sub
Sub Bloquear()
  FraDatos.Enabled = False
  BotTool 3, 17, ToolAuxiliares, False
End Sub
Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
        If Principal.ToolConsultas1.AbrirDevDatos("Digite ID", "Digite la identificacion del auxiliar", 2, 0) = True Then
          FufuSt = Principal.ToolConsultas1.DatSt
          If ExRecorset("Select IdAuxiliar from Auxiliares where IdAuxiliar='" & FufuSt & "'") = False Then
            Desbloquear
            limpiar
            TxtIdAuxiliar.Text = FufuSt
            TxtCO = Coperaciones
          Else
          MsgBox "Ya hay un auxiliar creado con esta identificacion, no se pueden crear dos datos con esta identificacion", vbCritical, "Este auxliar ya existe"
          End If
        End If
              
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            Bloquear
            AbrirRecorset rstUniversal, "Update Auxiliares NmAuxilar='" & TxtNmAuxiliar.Text & "', CO=" & Val(TxtCO) & ", TelAuxiliar='" & TxtTelefono & "' where IdAuxiliar=" & Val(TxtIdAuxiliar.Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          AbrirRecorset rstUniversal, "INSERT INTO Auxiliares VALUES ('" & TxtIdAuxiliar.Text & "','" & TxtNmAuxiliar.Text & "','" & TxtTelefono.Text & "'," & TxtCO.Text & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
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
      Principal.ToolConsultas1.AbrirDevConsulta 10, CnnPrincipal
      If Principal.ToolConsultas1.DatSt <> "" Then
        If BuscaRegistro("IdAuxiliar='" & Principal.ToolConsultas1.DatSt & "'", rstAuxiliares) = True Then Asignar
      End If
    Case 11 'Primero
      UPrimero rstAuxiliares
      Asignar
    Case 12 'Anterior
      UAnterior rstAuxiliares
      Asignar
    Case 13 'Siguiente
      USiguiente rstAuxiliares
      Asignar
    Case 14 'Ultimo
      UUltimo rstAuxiliares
      Asignar
    Case 16 'Cerrar
      CerrarRecorset rstAuxiliares
      Unload Me
    Case 17 'Actualizar
      rstAuxiliares.Requery
    Case 18 'Imprimir
  End Select
End Sub
Function Validacion() As Boolean
  If TxtNmAuxiliar.Text <> "" Then
    If TxtTelefono.Text <> "" Then
      Validacion = True
    Else
      Validacion = False: MsgTit "El auxiliar debe tener un telefono": TxtTelefono.SetFocus
    End If
  Else
    Validacion = False: MsgTit "El auxiliar debe tener un nombre": TxtNmAuxiliar.SetFocus
  End If
End Function
Private Sub ToolAuxiliares_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub
Private Sub TxtNmAuxiliar_GotFocus()
  EnfocarT TxtNmAuxiliar
End Sub
Private Sub TxtTelefono_GotFocus()
  EnfocarT TxtTelefono
End Sub
Private Sub TxtTelefono_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtTelefono, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
