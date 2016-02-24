VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAsesores 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asesores"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8055
      Begin VB.TextBox TxtCodigoInterface 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxtNmAsesor 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo interface:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label LblIdAsesor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   705
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "ID"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   165
      End
   End
   Begin MSComctlLib.Toolbar ToolAsesores 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9450
      _ExtentX        =   16669
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
Attribute VB_Name = "FrmAsesores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstAsesores As New ADODB.Recordset
Dim Editando As Boolean
Private Sub Form_Load()
  IconosTool ToolAsesores, Principal.IgListTool
  rstAsesores.CursorLocation = adUseServer
  AbrirRecorset rstAsesores, "SELECT asesores.* From asesores", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Asignar rstAsesores
End Sub
Private Sub Asignar(rstAsignar As ADODB.Recordset)
  LblIdAsesor.Caption = rstAsignar!IdAsesor
  TxtNmAsesor.Text = rstAsignar!NmAsesor & ""
  TxtCodigoInterface.Text = rstAsignar!codigo_interface & ""
End Sub
Private Sub limpiar()
  LblIdAsesor.Caption = ""
  TxtNmAsesor.Text = ""
  TxtCodigoInterface.Text = ""
End Sub

Private Sub Desbloquear()
  FraDatos.Enabled = True
  BotTool 3, 17, ToolAsesores, True
End Sub
Private Sub Bloquear()
  FraDatos.Enabled = False
  BotTool 3, 17, ToolAsesores, False
End Sub
Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
      Desbloquear
      limpiar
      Editando = False
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            Bloquear
            AbrirRecorset rstUniversal, "Update asesores set NmAsesor='" & TxtNmAsesor & "', codigo_interface='" & TxtCodigoInterface & "' where IdAsesor=" & Val(LblIdAsesor), CnnPrincipal, adOpenDynamic, adLockOptimistic
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          AbrirRecorset rstUniversal, "INSERT INTO asesores (NmAsesor, codigo_interface) VALUES ('" & TxtNmAsesor & "','" & TxtCodigoInterface & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
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
        Asignar rstAsesores
        Bloquear
      End If
    Case 9  'Buscar
      FrmBuscarAsesor.Show 1
      If FufuLo <> 0 Then
        AbrirRecorset rstUniversal, "Select*from asesores where IdAsesor=" & FufuLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          Asignar rstUniversal
        Else
          MsgBox "No se encontraron asesores con este codigo", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If
      
    Case 11 'Primero
      UPrimero rstAsesores
      Asignar rstAsesores
    Case 12 'Anterior
      UAnterior rstAsesores
      Asignar rstAsesores
    Case 13 'Siguiente
      USiguiente rstAsesores
      Asignar rstAsesores
    Case 14 'Ultimo
      UUltimo rstAsesores
      Asignar rstAsesores
    Case 16 'Cerrar
      CerrarRecorset rstAsesores
      FufuLo = Val(LblIdAsesor)
      'Principal.MnuManten.Enabled = True
      Unload Me
    Case 17 'Actualizar
      rstAsesores.Requery
    Case 18 'Imprimir
    Case 19
      MsgBox "No hay datos para actualizar"
  End Select
End Sub
Function Validacion() As Boolean
  If TxtNmAsesor.Text <> "" Then
    Validacion = True
  Else
    MsgBox "El asesor debe tener un nombre": Validacion = False: TxtNmAsesor.SetFocus
  End If
End Function
Private Sub ToolAsesores_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub
Private Sub TxtNmAsesor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub


