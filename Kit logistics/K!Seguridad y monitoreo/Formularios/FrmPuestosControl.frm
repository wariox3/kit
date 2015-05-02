VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPuestosControl 
   Caption         =   "Puesto de control"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9255
      Begin VB.TextBox TxtIdRuta 
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox TxtNmPuestoControl 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3000
         MaxLength       =   100
         TabIndex        =   2
         Top             =   240
         Width           =   5775
      End
      Begin VB.TextBox TxtIdPuestoControl 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LblNmRuta 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3000
         TabIndex        =   8
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Index           =   2
         Left            =   1800
         TabIndex        =   7
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Puesto control:"
         Height          =   195
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   1065
      End
   End
   Begin MSComctlLib.Toolbar ToolPuestoControl 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9630
      _ExtentX        =   16986
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
Attribute VB_Name = "FrmPuestosControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rstPuestosControl As New ADODB.Recordset
Dim Editando As Boolean

Private Sub Form_Load()
  IconosTool ToolPuestoControl, Principal.IgListTool
  rstPuestosControl.CursorLocation = adUseServer
  AbrirRecorset rstPuestosControl, "SELECT controlpost.*, rutas.NmRuta From controlpost LEFT JOIN rutas ON controlpost.IdRuta = rutas.IdRuta", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Asignar rstPuestosControl
End Sub
Private Sub Asignar(rstAsignar As ADODB.Recordset)
  TxtIdPuestoControl.Text = rstAsignar!IdControlPost
  TxtIdRuta.Text = rstAsignar!IdRuta & ""
  TxtNmPuestoControl.Text = rstAsignar!NmControlPost & ""
  LblNmRuta.Caption = rstAsignar!NmRuta & ""
End Sub

Private Sub limpiar()
  TxtIdPuestoControl.Text = ""
  TxtNmPuestoControl.Text = ""
  TxtIdRuta.Text = ""
  LblNmRuta.Caption = ""
End Sub
Private Sub Desbloquear()
  FraDatos.Enabled = True
  BotTool 3, 17, ToolPuestoControl, True
End Sub
Private Sub Bloquear()
  FraDatos.Enabled = False
  BotTool 3, 17, ToolPuestoControl, False
End Sub
Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
      Desbloquear
      limpiar
      
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            Bloquear
            AbrirRecorset rstUniversal, "Update controlpost set NmControlPost='" & TxtNmPuestoControl & "', IdRuta = " & Val(TxtIdRuta.Text) & " where IdControlPost=" & Val(TxtIdPuestoControl), CnnPrincipal, adOpenDynamic, adLockOptimistic
            AccionTool 17
            Asignar rstPuestosControl
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          AbrirRecorset rstUniversal, "INSERT INTO controlpost (NmControlPost, IdRuta) VALUES ('" & TxtNmPuestoControl & "', " & Val(TxtIdRuta.Text) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
          Bloquear
          AccionTool 17
          Asignar rstPuestosControl
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
        Asignar rstPuestosControl
        Bloquear
      End If
    Case 9  'Buscar
      If Principal.ToolConsultas1.AbrirConsultaGral("IdControlPost", "NmControlPost", "controlpost", CnnPrincipal) = True Then
        AbrirRecorset rstUniversal, "Select*from controlpost where IdControlPost=" & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          Asignar rstUniversal
        Else
          MsgBox "No se encontraron controlpost", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If

    Case 11 'Primero
      UPrimero rstPuestosControl
      Asignar rstPuestosControl
    Case 12 'Anterior
      UAnterior rstPuestosControl
      Asignar rstPuestosControl
    Case 13 'Siguiente
      USiguiente rstPuestosControl
      Asignar rstPuestosControl
    Case 14 'Ultimo
      UUltimo rstPuestosControl
      Asignar rstPuestosControl
    Case 16 'Cerrar
      CerrarRecorset rstPuestosControl
      'Principal.MnuManten.Enabled = True
      Unload Me
    Case 17 'Actualizar
      rstPuestosControl.Requery
    Case 18 'Imprimir
  End Select
End Sub
Function Validacion() As Boolean
  If TxtNmPuestoControl.Text <> "" Then
    If Val(TxtIdRuta.Text) <> 0 Then
      Validacion = True
    Else
      MsgBox "El punto de control debe tener una ruta": Validacion = False: TxtIdRuta.SetFocus
    End If
  Else
    MsgBox "El punto de control debe tener un nombre": Validacion = False: TxtNmPuestoControl.SetFocus
  End If
End Function
Private Sub ToolPuestoControl_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub

Private Sub TxtIdRuta_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    If Principal.ToolConsultas1.AbrirConsultaGral("IdRuta", "NmRuta", "rutas", CnnPrincipal) = True Then
      TxtIdRuta.Text = Principal.ToolConsultas1.DatLo
    End If
  End If
End Sub

Private Sub TxtNmPuestoControl_GotFocus()
  EnfocarT TxtNmPuestoControl
End Sub
Private Sub TxtNmPuestoControl_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
