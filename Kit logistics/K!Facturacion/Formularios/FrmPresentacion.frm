VERSION 5.00
Begin VB.Form FrmPresentacion 
   BorderStyle     =   0  'None
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmPresentacion.frx":0000
   ScaleHeight     =   4755
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label LblIdProducto 
      BackStyle       =   0  'Transparent
      Caption         =   "Producto"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   3120
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
   Begin VB.Label LblPropietario 
      BackStyle       =   0  'Transparent
      Caption         =   "Propietario"
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   2640
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
  Set CnnPrincipal = New ADODB.Connection
  Set rstUniversal = New ADODB.Recordset
  
  CnnPrincipal.CursorLocation = adUseClient
  rstUniversal.CursorLocation = adUseClient
  CnnPrincipal.Open "DRIVER=" & GetSetting("Kit Logistics", "Configuracion", "CnnDriver") & "; SERVER=" & GetSetting("Kit Logistics", "Configuracion", "CnnServidor") & "; PORT=" & GetSetting("Kit Logistics", "Configuracion", "CnnPuerto") & "; DATABASE=" & GetSetting("Kit Logistics", "Configuracion", "CnnBaseDatos") & "; PWD=" & GetSetting("Kit Logistics", "Configuracion", "CnnContraseña") & "; UID=" & GetSetting("Kit Logistics", "Configuracion", "CnnUsuario") & ";OPTION=3"
  Coperaciones = GetSetting("Kit Logistics", "Configuracion", "Coperaciones")
  Me.Show
    LblEmpresa = GetSetting("Kit Logistics", "InfoSoftware", "Empresa", "1")
    LblPropietario = GetSetting("Kit Logistics", "InfoSoftware", "Propietario", "1")
    LblIdProducto = GetSetting("Kit Logistics", "InfoSoftware", "Serial", "1")
    
    rstUniversal.Open "Select configuracion.* from configuracion", CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstUniversal.EOF = False Then
      If rstUniversal.Fields("fecha_vence_licencia") <= Date Then
        MsgBox "La licencia se encuentra vencida desde el " & rstUniversal.Fields("fecha_vence_licencia") & " por favor consulte al proveedor del software en su ciudad", vbCritical
        Unload Me
        Exit Sub
      End If
    End If
    rstUniversal.Close
    
  CodUsuarioActivo = IngresoSistema(CnnPrincipal, 3)
  If CodUsuarioActivo <> 0 Then Principal.Show
  Unload Me
End Sub



