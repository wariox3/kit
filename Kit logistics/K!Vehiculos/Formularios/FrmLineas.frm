VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmLineas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lineas..."
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7455
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   7215
      Begin VB.TextBox TxtCodigo 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtLinea 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   720
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TxtNmLinea 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   2
         Top             =   600
         Width           =   4335
      End
      Begin MSDataListLib.DataCombo CboMarcas 
         Height          =   315
         Left            =   2640
         TabIndex        =   0
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion:"
         Height          =   195
         Left            =   1680
         TabIndex        =   9
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Linea:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Marca:"
         Height          =   195
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
   End
   Begin MSComctlLib.Toolbar ToolLineas 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
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
Attribute VB_Name = "FrmLineas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstLineas As New ADODB.Recordset
Dim Editando As Boolean
Private Sub CboMarcas_GotFocus()
  'LlenarCombo CboMarcas, "NmMarca", "Marcas"
End Sub

Private Sub CboMarcas_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub CboMarcas_Validate(Cancel As Boolean)
  If CboMarcas.Text <> "" Then
    AbrirRecorset rstUniversal, "SELECT IdMarca, NmMarca From Marcas where NmMarca=" & CboMarcas.Text, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      CboMarcas.Tag = rstUniversal!IdMarca
    Else
      CboMarcas.Tag = 0: CboMarcas.Text = ""
    End If
    CerrarRecorset rstUniversal
  End If
End Sub


Private Sub Form_Load()
  IconosTool ToolLineas, Principal.IgListTool
  rstLineas.CursorLocation = adUseServer
  AbrirRecorset rstLineas, "SELECT*From Lineas", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Asignar
End Sub
Sub Asignar()
  TxtCodigo.Text = rstLineas!IdLinea
  CboMarcas.Tag = rstLineas!IdMarca
  TxtLinea.Text = rstLineas!Linea & ""
  TxtNmLinea.Text = rstLineas!NmLinea & ""
  CboMarcas.Text = ""
End Sub
Sub limpiar()
  TxtCodigo.Text = ""
  CboMarcas.Tag = ""
  TxtLinea.Text = ""
  TxtNmLinea.Text = ""
End Sub
Sub Desbloquear()
  FraDatos.Enabled = True
  BotTool 3, 17, ToolLineas, True
End Sub
Sub Bloquear()
  FraDatos.Enabled = False
  BotTool 3, 17, ToolLineas, False
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
            AbrirRecorset rstUniversal, "Update Lineas set IdMarca=" & Val(CboMarcas.Tag) & ", Linea='" & TxtLinea & "', NmLinea='" & TxtNmLinea & "' where IdLinea=" & Val(TxtCodigo.Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          AbrirRecorset rstUniversal, "INSERT INTO Lineas VALUES (" & Val(CboMarcas.Tag) & ",'" & TxtLinea.Text & "','" & TxtNmLinea.Text & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
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
      'FrmConsultaLineas.Show 1
      'If BuscaRegistro("IdLinea=" & FufuLo, rstLineas) = True Then Asignar
    Case 11 'Primero
      UPrimero rstLineas
      Asignar
    Case 12 'Anterior
      UAnterior rstLineas
      Asignar
    Case 13 'Siguiente
      USiguiente rstLineas
      Asignar
    Case 14 'Ultimo
      UUltimo rstLineas
      Asignar
    Case 16 'Cerrar
      CerrarRecorset rstLineas
      Unload Me
    Case 17 'Actualizar
      rstLineas.Requery
    Case 18 'Imprimir
    Case 19
      'If Val(CboMarcas.Tag) <> 0 Then CboMarcas.Text = DevResBus("SELECT IdMarca, NmMarca From Marcas where Idmarca=" & Val(CboMarcas.Tag), "NmMarca")
  End Select
End Sub
Function Validacion() As Boolean
  If TxtLinea.Text <> "" Then
    If TxtNmLinea.Text <> "" Then
      If Val(CboMarcas.Tag) <> 0 Then
        Validacion = True
      Else
        MsgBox "La linea debe tener una marca": Validacion = False: CboMarcas.SetFocus
      End If
      
    Else
      MsgBox "La linea debe tener una descripcion": Validacion = False: TxtNmLinea.SetFocus
    End If
  Else
    MsgBox "La linea debe tener una codificacion asignada por el ministerio de transporte": Validacion = False: TxtLinea.SetFocus
  End If
End Function
Private Sub ToolLineas_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub
Private Sub TxtLinea_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
