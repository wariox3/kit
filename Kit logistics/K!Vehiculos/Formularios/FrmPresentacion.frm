VERSION 5.00
Begin VB.Form FrmPresentacion 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmPresentacion.frx":0000
   ScaleHeight     =   4755
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label LblPropietario 
      BackStyle       =   0  'Transparent
      Caption         =   "Propietario"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label LblEmpresa 
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   2880
      Width           =   3615
   End
   Begin VB.Label LblIdProducto 
      BackStyle       =   0  'Transparent
      Caption         =   "Producto"
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   3120
      Width           =   3615
   End
End
Attribute VB_Name = "FrmPresentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim Msg As String
On Error GoTo SinConexion
  If App.PrevInstance Then
    Msg = App.EXEName & ".EXE" & " ya est� en ejecuci�n"
    MsgBox Msg, 16, "Aplicaci�n."
    End
  End If
  
  Me.Show
  Set CnnPrincipal = New ADODB.Connection
  Set rstUniversal = New ADODB.Recordset
      
  CnnPrincipal.CursorLocation = adUseClient
  rstUniversal.CursorLocation = adUseClient
  CnnPrincipal.Open "DRIVER=" & GetSetting("Kit Logistics", "Configuracion", "CnnDriver") & "; SERVER=" & GetSetting("Kit Logistics", "Configuracion", "CnnServidor") & "; PORT=" & GetSetting("Kit Logistics", "Configuracion", "CnnPuerto") & "; DATABASE=" & GetSetting("Kit Logistics", "Configuracion", "CnnBaseDatos") & "; PWD=" & GetSetting("Kit Logistics", "Configuracion", "CnnContrase�a") & "; UID=" & GetSetting("Kit Logistics", "Configuracion", "CnnUsuario") & ";OPTION=3"
  LblEmpresa = GetSetting("Kit Logistics", "InfoSoftware", "Empresa")
  LblPropietario = GetSetting("Kit Logistics", "InfoSoftware", "Propietario")
  LblIdProducto = GetSetting("Kit Logistics", "InfoSoftware", "Serial")
  CodUsuarioActivo = IngresoSistema(CnnPrincipal, 1)
  If CodUsuarioActivo <> 0 Then Principal.Show
  Unload Me
SinConexion:
  If Err.Number <> 0 Then
    MsgBox "No se ha podido conectar correctamente la base de datos BD: " & GetSetting("Kit Logistics", "Configuracion", "CnnBaseDatos") & " en Servidor:" & GetSetting("Kit Logistics", "Configuracion", "CnnServidor") & " Usuario:" & GetSetting("Kit Logistics", "Configuracion", "CnnUsuario") & ", este error puede ser a causa de la conexion; puede hacer lo siguiente:" & Chr(13) & "- Consulte al proveedor" & Chr(13) & "- Configure la conexion desde el menu herramientas y configurar la conexion con la BD en el modulo de administrador", vbCritical, "Error de conexion"
    Unload Me
  End If
End Sub
