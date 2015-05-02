VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm Principal 
   BackColor       =   &H8000000C&
   Caption         =   "Administrador 1.0.0.2 "
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9960
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicContenedor 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   9900
      TabIndex        =   0
      Top             =   0
      Width           =   9960
   End
   Begin MSComDlg.CommonDialog CDExa 
      Left            =   1320
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".kin"
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu MnuConcetarBD 
         Caption         =   "Conectar a BD"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "Opciones"
      Begin VB.Menu MnuConfiguracionGeneral 
         Caption         =   "Configuracion general"
      End
   End
   Begin VB.Menu MnuHerramientas 
      Caption         =   "Herramientas"
      Begin VB.Menu MnuControlarErrores 
         Caption         =   "Controlar errores de operacion del sistema"
      End
      Begin VB.Menu MnuEjecutarScripts 
         Caption         =   "Ejecutar scripts"
      End
      Begin VB.Menu MnuSep21 
         Caption         =   "-"
      End
      Begin VB.Menu MnuConfiguracionConBD 
         Caption         =   "Configuracion de conexion con la BD"
      End
   End
   Begin VB.Menu MnuCo 
      Caption         =   "Centros de operacion"
      Begin VB.Menu MnuIniciarCO 
         Caption         =   "Iniciar Manager"
      End
   End
   Begin VB.Menu MnuUsuarios 
      Caption         =   "Usuarios"
      Begin VB.Menu MnuMantenimientoUsuarios 
         Caption         =   "Mantenimiento"
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MnuConcetarBD_Click()
On Error GoTo SinConexion
  CnnPrincipal.Open "DRIVER=" & GetSetting("Kit Logistics", "Configuracion", "CnnDriver") & "; SERVER=" & GetSetting("Kit Logistics", "Configuracion", "CnnServidor") & ";  PORT=" & GetSetting("Kit Logistics", "Configuracion", "CnnPuerto") & "; DATABASE=" & GetSetting("Kit Logistics", "Configuracion", "CnnBaseDatos") & "; PWD=" & GetSetting("Kit Logistics", "Configuracion", "CnnContraseña") & "; UID=" & GetSetting("Kit Logistics", "Configuracion", "CnnUsuario") & ";OPTION=3"
  MsgBox "Base de datos conectada con exito", vbInformation
  MnuConcetarBD.Enabled = False
  ActivarOptConexiones
SinConexion:
  If Err.Number <> 0 Then
    MsgBox "No se ha podido conectar correctamente la base de datos, este error puede ser a causa de la conexion; puede hacer lo siguiente:" & Chr(13) & "- Consulte al proveedor" & Chr(13) & "- Configure la conexion desde el menu herramientas y configurar la conexion con la BD" & Chr(13) & "- Trabaje sin los modulos que requieren conexion", vbCritical, "Error de conexion"
  End If
End Sub
Private Sub MnuConfiguracionConBD_Click()
  FrmConectarBD.Show 1
End Sub

Private Sub MnuConfiguracionGeneral_Click()
  FrmConfiguracionGeneral.Show 1
End Sub

Private Sub MnuEjecutarScripts_Click()
  FrmEjecutarScripts.Show 1
End Sub

Private Sub MnuIniciarCO_Click()
  If CnnPrincipal.State = adStateOpen Then
    If FormAbierto = False Then
      FormAbierto = True
      FrmCentrosOperaciones.Show
    End If
  Else
    MsgBox "Debe conectarse primero a la base de datos", vbCritical, "Sin conexion a la base de datos"
  End If
  
End Sub

Private Sub MnuMantenimientoUsuarios_Click()
  If CnnPrincipal.State = adStateOpen Then
    If FormAbierto = False Then
      FormAbierto = True
      FrmUsuarios.Show
    End If
  Else
    MsgBox "Debe conectarse primero a la base de datos", vbCritical, "Sin conexion a la base de datos"
  End If
End Sub

Private Sub MnuSalir_Click()
  Unload Me
End Sub

Private Sub ActivarOptConexiones()
  If MnuConcetarBD.Enabled = False Then
    MnuCo.Enabled = True
    MnuUsuarios.Enabled = True
  End If
End Sub
