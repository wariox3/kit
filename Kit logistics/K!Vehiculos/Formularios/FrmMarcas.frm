VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMarcas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Marcas..."
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7440
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   7215
      Begin VB.TextBox TxtIdMarca 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtNmMarca 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2400
         TabIndex        =   0
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox TxtMinTransporte 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2400
         MaxLength       =   8
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Codigo MinTransporte:"
         Height          =   195
         Left            =   720
         TabIndex        =   7
         Top             =   600
         Width           =   1605
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Marca:"
         Height          =   195
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   540
      End
   End
   Begin MSComctlLib.Toolbar ToolMarcas 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7440
      _ExtentX        =   13123
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
            Object.ToolTipText     =   "Guarda la informacio [F10]"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Object.ToolTipText     =   "Editar la informacion guardada [F11]"
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
Attribute VB_Name = "FrmMarcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstMarcas As New ADODB.Recordset
Dim Editando As Boolean

Private Sub Form_Load()
  IconosTool ToolMarcas, Principal.IgListTool
  rstMarcas.CursorLocation = adUseServer
  AbrirRecorset rstMarcas, "SELECT*From Marcas", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Asignar
End Sub
Sub Asignar()
  TxtIdMarca.Text = rstMarcas!IdMarca
  TxtNmMarca.Text = rstMarcas!NmMarca & ""
  TxtMinTransporte.Text = rstMarcas!CodMinTrans & ""
End Sub
Sub limpiar()
  TxtIdMarca.Text = ""
  TxtNmMarca.Text = ""
  TxtMinTransporte.Text = ""
End Sub
Sub Desbloquear()
  FraDatos.Enabled = True
  BotTool 3, 17, ToolMarcas, True
End Sub
Sub Bloquear()
  FraDatos.Enabled = False
  BotTool 3, 17, ToolMarcas, False
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
            AbrirRecorset rstUniversal, "Update Marcas set NmMarca='" & TxtNmMarca & "', CodMinTrans='" & TxtMinTransporte & "' where IdDepartamento=" & Val(TxtIdMarca), CnnPrincipal, adOpenDynamic, adLockOptimistic
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          AbrirRecorset rstUniversal, "INSERT INTO Marcas VALUES ('" & TxtNmMarca & "','" & TxtMinTransporte & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
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
      'FrmConsultaMarcas.Show 1
      'If BuscaRegistro("IdMarca=" & FufuLo, rstMarcas) = True Then Asignar
    Case 11 'Primero
      UPrimero rstMarcas
      Asignar
    Case 12 'Anterior
      UAnterior rstMarcas
      Asignar
    Case 13 'Siguiente
      USiguiente rstMarcas
      Asignar
    Case 14 'Ultimo
      UUltimo rstMarcas
      Asignar
    Case 16 'Cerrar
      CerrarRecorset rstMarcas
      Unload Me
    Case 17 'Actualizar
      rstMarcas.Requery
    Case 18 'Imprimir
  End Select
End Sub
Function Validacion() As Boolean
  If TxtNmMarca.Text <> "" Then
    If TxtMinTransporte.Text <> "" Then
      Validacion = True
    Else
      MsgBox "La carroceria debe tener un codigo del ministerio": Validacion = False: TxtMinTransporte.SetFocus
    End If
  Else
    MsgBox "La carroceria debe tener un nombre": Validacion = False: TxtNmMarca.SetFocus
  End If
End Function
Private Sub ToolMarcas_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub
Private Sub TxtMinTransporte_GotFocus()
  EnfocarT TxtMinTransporte
End Sub
Private Sub TxtMinTransporte_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtMinTransporte, KeyAscii, 1
End Sub
Private Sub TxtNmMarca_GotFocus()
  EnfocarT TxtNmMarca
End Sub
Private Sub TxtNmMarca_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
