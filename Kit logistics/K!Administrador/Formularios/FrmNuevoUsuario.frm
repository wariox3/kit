VERSION 5.00
Begin VB.Form FrmNuevoUsuario 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nuevo Usuario..."
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtRemiteMail 
      Height          =   285
      Left            =   1080
      TabIndex        =   20
      Top             =   2520
      Width           =   4815
   End
   Begin VB.TextBox TxtClaveMail 
      Height          =   285
      Left            =   1080
      TabIndex        =   19
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox TxtUsuarioMail 
      Height          =   285
      Left            =   1080
      TabIndex        =   18
      Top             =   1800
      Width           =   4815
   End
   Begin VB.CheckBox ChkInactivo 
      Caption         =   "Inactivo"
      Height          =   255
      Left            =   4440
      TabIndex        =   14
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox ChkDatosBasicos 
      Caption         =   "Datos Basico"
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CheckBox ChkVehiculos 
      Caption         =   "Vehiculos"
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CheckBox ChkRecogidas 
      Caption         =   "Recogidas"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   840
      Width           =   1215
   End
   Begin VB.CheckBox ChkFacturacion 
      Caption         =   "Facturacion"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox ChkMonitoreo 
      Caption         =   "Monitoreo"
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CheckBox ChkMovimiento 
      Caption         =   "Movimiento"
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox TxtReContraseña 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4440
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox TxtContraseña 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox TxtNombre 
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Remite Mail:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Clave Mail:"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario Mail:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   915
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Re-Contraseña:"
      Height          =   195
      Index           =   2
      Left            =   3240
      TabIndex        =   6
      Top             =   480
      Width           =   1110
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   600
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña:"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "FrmNuevoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
  If TxtNombre <> "" Then
    If TxtContraseña.Text <> "" Then
      If Len(TxtContraseña) > 5 Then
        If TxtContraseña.Text = TxtReContraseña.Text Then
          If II = 1 Then
            AbrirRecorset rstUniversal, "Insert into Usuarios (NmUsuario, Contraseña, ModMovimiento, ModMonitoreo, ModFacturacion, ModRecogidas, ModVehiculos, ModDatosBasicos, Inactivo) values ('" & TxtNombre.Text & "', '" & EncryptString("mario", TxtContraseña.Text, 1) & "', " & ChkMovimiento.Value & ", " & ChkMonitoreo.Value & ", " & ChkFacturacion.Value & ", " & ChkRecogidas.Value & ", " & ChkVehiculos.Value & ", " & ChkDatosBasicos.Value & ", " & ChkInactivo.Value & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
            MsgBox "Usuario creado con exito", vbInformation
            Unload Me
          Else
            AbrirRecorset rstUniversal, "Update Usuarios Set NmUsuario='" & TxtNombre.Text & "', Contraseña='" & EncryptString("mario", TxtContraseña.Text, 1) & "', ModMovimiento=" & ChkMovimiento.Value & ", ModMonitoreo=" & ChkMonitoreo.Value & ", ModFacturacion=" & ChkFacturacion.Value & ", ModRecogidas=" & ChkRecogidas.Value & ", ModVehiculos=" & ChkVehiculos.Value & ", ModDatosBasicos=" & ChkDatosBasicos.Value & ", Inactivo=" & ChkInactivo.Value & ", UsuarioMail='" & TxtUsuarioMail.Text & "', ClaveMail='" & TxtClaveMail.Text & "', RemiteMail='" & TxtRemiteMail.Text & "'  where IdUsuario=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
            MsgBox "Usuario actualizado con exito", vbInformation
            Unload Me
          End If
        Else
          MsgBox "las contraseñas deben ser iguales", vbCritical
          TxtContraseña.SetFocus
        End If
      Else
        MsgBox "La contraseña debe ser minimo de 6 digitos", vbCritical
        TxtContraseña.SetFocus
      End If
    Else
      MsgBox "Debe digitar una contraseña", vbCritical
      TxtContraseña.SetFocus
    End If
  Else
    MsgBox "El nombre no puede estar vacío", vbCritical
    TxtNombre.SetFocus
  End If
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  If II = 2 Then
    Me.Caption = "Editar Usuario..."
    AbrirRecorset rstUniversal, "Select usuarios.* from usuarios where IDUsuario=" & FufuLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      TxtNombre.Text = rstUniversal!NmUsuario
      TxtContraseña.Text = EncryptString("mario", rstUniversal!Contraseña, 2)
      TxtReContraseña.Text = EncryptString("mario", rstUniversal!Contraseña, 2)
      ChkMovimiento.Value = DevCheck(rstUniversal!ModMovimiento)
      ChkMonitoreo.Value = DevCheck(rstUniversal!ModMonitoreo)
      ChkFacturacion.Value = DevCheck(rstUniversal!ModFacturacion)
      ChkRecogidas.Value = DevCheck(rstUniversal!ModRecogidas)
      ChkVehiculos.Value = DevCheck(rstUniversal!ModVehiculos)
      ChkDatosBasicos.Value = DevCheck(rstUniversal!ModDatosBasicos)
      ChkInactivo.Value = DevCheck(rstUniversal!Inactivo)
      TxtUsuarioMail.Text = rstUniversal.Fields("UsuarioMail") & ""
      TxtClaveMail.Text = rstUniversal.Fields("ClaveMail") & ""
      TxtRemiteMail.Text = rstUniversal.Fields("RemiteMail") & ""
    CerrarRecorset rstUniversal
  End If
End Sub

Private Sub LblTitulo_DblClick(Index As Integer)
  If InputBox("Digite el codigo de administrador") = "0313" Then
    TxtReContraseña.PasswordChar = ""
  End If
End Sub

Private Sub TxtContraseña_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub TxtReContraseña_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
