VERSION 5.00
Begin VB.Form FrmPresentacion 
   BorderStyle     =   0  'None
   Caption         =   "Ingreso [Recogidas]"
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   HelpContextID   =   2
   Icon            =   "FrmPresentacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmPresentacion.frx":0CCA
   ScaleHeight     =   4755
   ScaleWidth      =   7320
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
  Me.Show
  Set CnnPrincipal = New ADODB.Connection
  Set rstUniversal = New ADODB.Recordset
  Set rstUniversalAux = New ADODB.Recordset

      
  CnnPrincipal.CursorLocation = adUseClient
  rstUniversal.CursorLocation = adUseClient
  rstUniversalAux.CursorLocation = adUseClient

      
  RutaLocal = GetSetting("Kit Logistics", "Configuracion", "RutaLocal")
      
  If CpExisteFichero(GetSetting("Kit Logistics", "Recogidas", "InfRecogidas")) = True Then
    ArchivoInf = GetSetting("Kit Logistics", "Recogidas", "InfRecogidas")
    
    Coperaciones = Val(GetSetting("Kit Logistics", "Configuracion", "Coperaciones"))
    CnnPrincipal.Open "DRIVER=" & GetSetting("Kit Logistics", "Configuracion", "CnnDriver") & "; SERVER=" & GetSetting("Kit Logistics", "Configuracion", "CnnServidor") & "; PORT=" & GetSetting("Kit Logistics", "Configuracion", "CnnPuerto") & "; DATABASE=" & GetSetting("Kit Logistics", "Configuracion", "CnnBaseDatos") & "; PWD=" & GetSetting("Kit Logistics", "Configuracion", "CnnContraseña") & "; UID=" & GetSetting("Kit Logistics", "Configuracion", "CnnUsuario") & ";OPTION=3"
    
    LblEmpresa = GetSetting("Kit Logistics", "InfoSoftware", "Empresa")
    LblPropietario = GetSetting("Kit Logistics", "InfoSoftware", "Propietario")
    LblIdProducto = GetSetting("Kit Logistics", "InfoSoftware", "Serial")
    
    CodUsuarioActivo = IngresoSistema(CnnPrincipal, 4)
      If CodUsuarioActivo <> 0 Then Principal.Show
  Else
    MsgBox "No existe el archivo de informacion, verifique el administrador", vbCritical, "No se puede iniciar el sistema"
  End If
  Unload Me
SinConexion:
  If Err.Number <> 0 Then
    MsgBox "No se ha podido conectar correctamente la base de datos BD: " & GetSetting("Kit Logistics", "Configuracion", "CnnBaseDatos") & " en Servidor:" & GetSetting("Kit Logistics", "Configuracion", "CnnServidor") & " Usuario:" & GetSetting("Kit Logistics", "Configuracion", "CnnUsuario") & ", este error puede ser a causa de la conexion; puede hacer lo siguiente:" & Chr(13) & "- Consulte al proveedor" & Chr(13) & "- Configure la conexion desde el menu herramientas y configurar la conexion con la BD en el modulo de administrador" & Chr(13) & Err.Description, vbCritical, "Error de conexion"
    Unload Me
  End If
End Sub
