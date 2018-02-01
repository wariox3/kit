VERSION 5.00
Begin VB.Form FrmContraseña 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control de acceso [Administrador]"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkSinConexion 
      Caption         =   "Ingresar sin conexion"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton CmdConfigurar 
      Caption         =   "Configurar conexion servidor"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton CmdIngresar 
      Caption         =   "Ingresar"
      Default         =   -1  'True
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox TxtContraseña 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña de administrador:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FrmContraseña"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub CmdConfigurar_Click()
  Dim Clave As String
  Clave = InputBox("Digite la clave:", "Clave admin")
  If Clave = "850903..." Then
    FrmConectarBD.Show 1
  End If
End Sub

Private Sub CmdIngresar_Click()
On Error GoTo SinConexion
    If TxtContraseña.Text = GetSetting("Kit Logistics", "Configuracion", "contrasena") Then
      If ChkSinConexion.Value = 0 Then
        CnnPrincipal.Open "DRIVER=" & GetSetting("Kit Logistics", "Configuracion", "CnnDriver") & "; SERVER=" & GetSetting("Kit Logistics", "Configuracion", "CnnServidor") & "; PORT=" & GetSetting("Kit Logistics", "Configuracion", "CnnPuerto") & "; DATABASE=" & GetSetting("Kit Logistics", "Configuracion", "CnnBaseDatos") & "; PWD=" & GetSetting("Kit Logistics", "Configuracion", "CnnContraseña") & "; UID=" & GetSetting("Kit Logistics", "Configuracion", "CnnUsuario") & ";OPTION=3"
        Principal.MnuConcetarBD.Enabled = False
        Principal.Show
        Unload Me
      Else
        Principal.Show
        Unload Me
      End If
    Else
      MsgBox "Contraseña incorrecta", vbCritical, "Contraseña incorrecta"
    End If
SinConexion:
  If Err.Number <> 0 Then
    MsgBox "No se ha podido conectar correctamente la base de datos, este error puede ser a causa de la conexion con la base de datos; puede hacer lo siguiente:" & Chr(13) & "- Consulte al proveedor" & Chr(13) & "- Configure la conexion desde el boton configurar" & Chr(13) & "- Ingrese sin conexion", vbCritical, "Error de conexion"
  End If
End Sub

Private Sub Form_Load()
  CnnPrincipal.CursorLocation = adUseClient
  rstUniversal.CursorLocation = adUseClient
End Sub
