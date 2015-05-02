VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCiudades 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ciudades..."
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10125
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   8055
      Begin VB.TextBox TxtCodigoMunicipio 
         Height          =   285
         Left            =   6120
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TxtCuentaCartera 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   6
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox TxtCuentaManejo 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox TxtCuentaFlete 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CheckBox ChkReexpedicion 
         Caption         =   "Reexpedicion"
         Height          =   255
         Left            =   6360
         TabIndex        =   14
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox TxtMinTransporte 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TxtIdDepartamento 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox TxtNmCiudad 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Codigo municipio:"
         Height          =   195
         Left            =   4800
         TabIndex        =   18
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cartera:"
         Height          =   195
         Left            =   435
         TabIndex        =   17
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta manejo:"
         Height          =   195
         Left            =   420
         TabIndex        =   16
         Top             =   1680
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta flete:"
         Height          =   195
         Left            =   630
         TabIndex        =   15
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Divipol:"
         Height          =   195
         Index           =   3
         Left            =   465
         TabIndex        =   12
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Departamento:"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   11
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label LblNmDepartamento 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "ID"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   165
      End
      Begin VB.Label LblIdCiudad 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   705
      End
   End
   Begin MSComctlLib.Toolbar ToolCiudades 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10125
      _ExtentX        =   17859
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
Attribute VB_Name = "FrmCiudades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCiudades As New ADODB.Recordset
Dim Editando As Boolean
Private Sub Form_Load()
  IconosTool ToolCiudades, Principal.IgListTool
  rstCiudades.CursorLocation = adUseServer
  AbrirRecorset rstCiudades, "SELECT ciudades.* From ciudades", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Asignar rstCiudades
End Sub
Private Sub Asignar(rstAsignar As ADODB.Recordset)
  LblIdCiudad.Caption = rstAsignar!IdCiudad
  TxtIdDepartamento.Text = rstAsignar!IdDepartamento
  TxtNmCiudad.Text = rstAsignar!NmCiudad & ""
  TxtMinTransporte.Text = rstAsignar!CodigoDivision & ""
  TxtCodigoMunicipio.Text = rstAsignar!CodigoMunicipio & ""
  TxtCuentaFlete.Text = rstAsignar!CuentaFlete & ""
  TxtCuentaManejo.Text = rstAsignar!CuentaManejo & ""
  TxtCuentaCartera.Text = rstAsignar!CuentaCartera & ""
  ChkReexpedicion.value = DevCheck(rstAsignar!Reexpedicion)
  LimpiarConsulta
End Sub
Private Sub limpiar()
  LblIdCiudad.Caption = ""
  TxtNmCiudad.Text = ""
  TxtIdDepartamento.Text = ""
  TxtMinTransporte.Text = ""
  TxtCodigoMunicipio.Text = ""
  TxtCuentaFlete.Text = ""
  TxtCuentaManejo.Text = ""
  ChkReexpedicion.value = 0
End Sub
Private Sub LimpiarConsulta()
  LblNmDepartamento.Caption = ""
End Sub
Private Sub Desbloquear()
  FraDatos.Enabled = True
  BotTool 3, 17, ToolCiudades, True
End Sub
Private Sub Bloquear()
  FraDatos.Enabled = False
  BotTool 3, 17, ToolCiudades, False
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
            AbrirRecorset rstUniversal, "Update Ciudades set NmCiudad='" & TxtNmCiudad & "', IdDepartamento=" & Val(TxtIdDepartamento) & ", CodigoDivision='" & TxtMinTransporte & "', CodigoMunicipio = '" & TxtCodigoMunicipio.Text & "' , Reexpedicion = " & ChkReexpedicion.value & ", CuentaFlete='" & TxtCuentaFlete & "', CuentaManejo='" & TxtCuentaManejo.Text & "', CuentaCartera='" & TxtCuentaCartera.Text & "' where IdCiudad=" & Val(LblIdCiudad), CnnPrincipal, adOpenDynamic, adLockOptimistic
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          AbrirRecorset rstUniversal, "INSERT INTO Ciudades (NmCiudad, IdDepartamento, CodigoDivision, Reexpedicion, CuentaFlete, CuentaManejo, CuentaCartera, CodigoMunicipio) VALUES ('" & TxtNmCiudad & "', " & Val(TxtIdDepartamento) & ", '" & TxtMinTransporte & "', " & ChkReexpedicion.value & ", '" & TxtCuentaFlete.Text & "', '" & TxtCuentaManejo.Text & "', '" & TxtCuentaCartera.Text & "', '" & TxtCodigoMunicipio.Text & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
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
        Asignar rstCiudades
        Bloquear
      End If
    Case 9  'Buscar
      If Principal.ToolConsultas1.AbrirConsultaGral("IdCiudad", "NmCiudad", "Ciudades", CnnPrincipal) = True Then
        AbrirRecorset rstUniversal, "Select*from Ciudades where IdCiudad=" & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          Asignar rstUniversal
        Else
          MsgBox "No se encontraron rutas con este codigo", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If
    Case 11 'Primero
      UPrimero rstCiudades
      Asignar rstCiudades
    Case 12 'Anterior
      UAnterior rstCiudades
      Asignar rstCiudades
    Case 13 'Siguiente
      USiguiente rstCiudades
      Asignar rstCiudades
    Case 14 'Ultimo
      UUltimo rstCiudades
      Asignar rstCiudades
    Case 16 'Cerrar
      CerrarRecorset rstCiudades
      FufuLo = Val(LblIdCiudad)
      'Principal.MnuManten.Enabled = True
      Unload Me
    Case 17 'Actualizar
      rstCiudades.Requery
    Case 18 'Imprimir
    Case 19
      If Val(TxtIdDepartamento) <> 0 Then LblNmDepartamento.Caption = DevResBus("SELECT IdDepartamento, NmDepartamento From Departamentos where IdDepartamento=" & TxtIdDepartamento.Text, "NmDepartamento", CnnPrincipal)
  End Select
End Sub
Function Validacion() As Boolean
  If TxtNmCiudad.Text <> "" Then
    If Val(TxtIdDepartamento) <> 0 Then
      If TxtMinTransporte.Text <> "" Then
        Validacion = True
      Else
        MsgBox "La ciudad debe tener un codigo del ministerio": Validacion = False: TxtMinTransporte.SetFocus
      End If
    Else
      MsgBox "La ciudad debe tener un departamento": Validacion = False: TxtIdDepartamento.SetFocus
    End If
  Else
    MsgBox "La ciudad debe tener un nombre": Validacion = False: TxtNmCiudad.SetFocus
  End If
End Function



Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtCuentaManejo, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub ToolCiudades_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub

Private Sub TxtCodigoMunicipio_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtCodigoMunicipio, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtCuentaFlete_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtCuentaFlete, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtCuentaManejo_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtCuentaManejo, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    If Principal.ToolConsultas1.AbrirConsultaGral("IdDepartamento", "NmDepartamento", "Departamentos", CnnPrincipal) = True Then
      TxtIdDepartamento.Text = Principal.ToolConsultas1.DatLo
    End If
  End If
End Sub

Private Sub TxtIdDepartamento_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtIdDepartamento, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdDepartamento_Validate(Cancel As Boolean)
  If Val(TxtIdDepartamento) <> 0 Then
    AbrirRecorset rstUniversal, "SELECT IdDepartamento, NmDepartamento FROM Departamentos where IdDepartamento=" & Val(TxtIdDepartamento), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      LblNmDepartamento = rstUniversal!NmDepartamento & ""
    Else
      LblNmDepartamento = "": TxtIdDepartamento = ""
    End If
    CerrarRecorset rstUniversal
  End If
End Sub

Private Sub TxtMinTransporte_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtMinTransporte, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub TxtNmCiudad_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
