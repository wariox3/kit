VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmRutasUrbanas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rutas Urbanas"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9975
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8415
      Begin VB.TextBox TxtCO 
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TxtIdRutaUrbana 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtNmRutaUrbana 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   1
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label LblConsulta 
         AutoSize        =   -1  'True
         Caption         =   "CO:"
         Height          =   195
         Index           =   2
         Left            =   7320
         TabIndex        =   6
         Top             =   240
         Width           =   270
      End
      Begin VB.Label LblConsulta 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   600
      End
      Begin VB.Label LblConsulta 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   540
      End
   End
   Begin MSComctlLib.Toolbar ToolRutasUrbanas 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
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
Attribute VB_Name = "FrmRutasUrbanas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstRutasUrbanas As New ADODB.Recordset
Dim Editando As Boolean
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  LaTecla (KeyCode), ToolRutasUrbanas
End Sub
Private Sub Form_Load()
  IconosTool ToolRutasUrbanas, Principal.IgListTool
  rstRutasUrbanas.CursorLocation = adUseServer
  AbrirRecorset rstRutasUrbanas, "SELECT*From RutasUrbanas", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Asignar
End Sub
Sub Asignar()
  TxtIdRutaUrbana.Text = rstRutasUrbanas!IdRutaRec
  TxtNmRutaUrbana.Text = rstRutasUrbanas!NmRuta & ""
  TxtCO.Text = rstRutasUrbanas!CO & ""
End Sub
Sub limpiar()
  TxtIdRutaUrbana.Text = ""
  TxtNmRutaUrbana.Text = ""
  TxtCO.Text = ""
End Sub
Sub Desbloquear()
  FraDatos.Enabled = True
  BotTool 3, 17, ToolRutasUrbanas, True
End Sub
Sub Bloquear()
  FraDatos.Enabled = False
  BotTool 3, 17, ToolRutasUrbanas, False
End Sub
Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
      Desbloquear
      limpiar
      TxtCO = Coperaciones
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            Bloquear
            AbrirRecorset rstUniversal, "Update RutasUrbanas set NmRuta='" & TxtNmRutaUrbana.Text & "', CO=" & Val(TxtCO) & " where IdRutaRec=" & Val(TxtIdRutaUrbana.Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          AbrirRecorset rstUniversal, "INSERT INTO rutasurbanas (NmRuta, CO) VALUES ('" & TxtNmRutaUrbana.Text & "'," & TxtCO & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
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
      Principal.ToolConsultas1.AbrirDevConsultaCO 1, Coperaciones, CnnPrincipal
      If Principal.ToolConsultas1.DatLo <> 0 Then
        If BuscaRegistro("IdRutaRec=" & Principal.ToolConsultas1.DatLo, rstRutasUrbanas) = True Then Asignar
      End If
    Case 11 'Primero
      UPrimero rstRutasUrbanas
      Asignar
    Case 12 'Anterior
      UAnterior rstRutasUrbanas
      Asignar
    Case 13 'Siguiente
      USiguiente rstRutasUrbanas
      Asignar
    Case 14 'Ultimo
      UUltimo rstRutasUrbanas
      Asignar
    Case 16 'Cerrar
      CerrarRecorset rstRutasUrbanas
      Unload Me
    Case 17 'Actualizar
      rstRutasUrbanas.Requery
    Case 18 'Imprimir
  End Select
End Sub
Function Validacion() As Boolean
   If TxtNmRutaUrbana.Text <> "" Then
    Validacion = True
   Else
    Validacion = False: MsgTit "La ruta debe tener un nombre": TxtNmRutaUrbana.SetFocus
   End If
End Function
Private Sub ToolRutasUrbanas_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub
Private Sub TxtNmRutaUrbana_GotFocus()
  EnfocarT TxtNmRutaUrbana
End Sub
