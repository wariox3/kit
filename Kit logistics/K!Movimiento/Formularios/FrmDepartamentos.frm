VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDepartamentos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Departamentos..."
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9180
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8055
      Begin VB.TextBox TxtMinTransporte 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TxtNmDepartamento 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2760
         MaxLength       =   100
         TabIndex        =   2
         Top             =   240
         Width           =   5175
      End
      Begin VB.TextBox TxtIdDepartamento 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Departamento:"
         Height          =   195
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Codigo MinTransporte:"
         Height          =   195
         Left            =   1080
         TabIndex        =   4
         Top             =   600
         Width           =   1605
      End
   End
   Begin MSComctlLib.Toolbar ToolDepartamentos 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
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
Attribute VB_Name = "FrmDepartamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstDepartamentos As New ADODB.Recordset
Dim Editando As Boolean

Private Sub Form_Load()
  IconosTool ToolDepartamentos, Principal.IgListTool
  rstDepartamentos.CursorLocation = adUseServer
  AbrirRecorset rstDepartamentos, "SELECT*From Departamentos", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Asignar rstDepartamentos
End Sub
Private Sub Asignar(rstAsignar As ADODB.Recordset)
  TxtIdDepartamento.Text = rstDepartamentos!IdDepartamento
  TxtNmDepartamento.Text = rstDepartamentos!NmDepartamento & ""
  TxtMinTransporte.Text = rstDepartamentos!CodMinTrans & ""
End Sub
Private Sub limpiar()
  TxtIdDepartamento.Text = ""
  TxtNmDepartamento.Text = ""
  TxtMinTransporte.Text = ""
End Sub
Private Sub Desbloquear()
  FraDatos.Enabled = True
  BotTool 3, 17, ToolDepartamentos, True
End Sub
Private Sub Bloquear()
  FraDatos.Enabled = False
  BotTool 3, 17, ToolDepartamentos, False
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
            AbrirRecorset rstUniversal, "Update Departamentos set NmDepartamento='" & TxtNmDepartamento & "', CodMinTrans='" & TxtMinTransporte & "' where IdDepartamento=" & Val(TxtIdDepartamento), CnnPrincipal, adOpenDynamic, adLockOptimistic
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          AbrirRecorset rstUniversal, "INSERT INTO Departamentos (NmDepartamento, CodMinTrans) VALUES ('" & TxtNmDepartamento & "','" & TxtMinTransporte & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
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
        Asignar rstDepartamentos
        Bloquear
      End If
    Case 9  'Buscar
      If Principal.ToolConsultas1.AbrirConsultaGral("IdDepartamento", "NmDepartamento", "Departamentos", CnnPrincipal) = True Then
        AbrirRecorset rstUniversal, "Select*from Departamentos where IdDepartamento=" & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          Asignar rstUniversal
        Else
          MsgBox "No se encontraron rutas con este codigo", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If
    Case 11 'Primero
      UPrimero rstDepartamentos
      Asignar rstDepartamentos
    Case 12 'Anterior
      UAnterior rstDepartamentos
      Asignar rstDepartamentos
    Case 13 'Siguiente
      USiguiente rstDepartamentos
      Asignar rstDepartamentos
    Case 14 'Ultimo
      UUltimo rstDepartamentos
      Asignar rstDepartamentos
    Case 16 'Cerrar
      CerrarRecorset rstDepartamentos
      'Principal.MnuManten.Enabled = True
      Unload Me
    Case 17 'Actualizar
      rstDepartamentos.Requery
    Case 18 'Imprimir
  End Select
End Sub
Function Validacion() As Boolean
  If TxtNmDepartamento.Text <> "" Then
    If TxtMinTransporte.Text <> "" Then
      Validacion = True
    Else
      MsgBox "El departamento debe tener un codigo del ministerio": Validacion = False: TxtNmDepartamento.SetFocus
    End If
  Else
    MsgBox "El departamento debe tener un nombre": Validacion = False: TxtNmDepartamento.SetFocus
  End If
End Function

Private Sub ToolDepartamentos_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub

Private Sub TxtMinTransporte_GotFocus()
  EnfocarT TxtMinTransporte
End Sub

Private Sub TxtMinTransporte_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtMinTransporte, KeyAscii, 1
End Sub

Private Sub TxtNmDepartamento_GotFocus()
  EnfocarT TxtNmDepartamento
End Sub

Private Sub TxtNmDepartamento_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
