VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8072DC64-8993-404F-8876-E5392C16A5C4}#1.0#0"; "PyConsultasKL.ocx"
Begin VB.MDIForm Principal 
   BackColor       =   &H8000000C&
   Caption         =   "Que!elp - Editor de listas de precios 1.0"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   -1875
   ClientWidth     =   12840
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin ConsultasKL.ToolConsultas ToolConsultas1 
      Left            =   6240
      Top             =   4440
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin MSComDlg.CommonDialog CDExa 
      Left            =   120
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   12780
      TabIndex        =   0
      Top             =   8985
      Width           =   12840
      Begin VB.TextBox TxtMensaje 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2895
      End
      Begin MSComctlLib.ProgressBar PgsPrincipal 
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Min             =   1e-4
      End
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu MnuConectarBaseDatos 
         Caption         =   "Conectar con la base de datos"
      End
      Begin VB.Menu MnuConectarAlIniciar 
         Caption         =   "Conectar al iniciar"
      End
      Begin VB.Menu MnuSep24 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAbrir 
         Caption         =   "Abrir lista de precios de archivo"
         Enabled         =   0   'False
         Shortcut        =   ^A
      End
      Begin VB.Menu MnuAbrirPreciosBD 
         Caption         =   "Abrir lista de precios de la Base de datos"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuSep23 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExportar 
         Caption         =   "Exportar"
         Enabled         =   0   'False
         Begin VB.Menu MnuExportExcel 
            Caption         =   "A Excel"
         End
         Begin VB.Menu MnuATexto 
            Caption         =   "A Texto"
         End
         Begin VB.Menu MnuExportAhtml 
            Caption         =   "A HTML"
         End
      End
      Begin VB.Menu MnuImportar 
         Caption         =   "Importar"
         Begin VB.Menu MnuImportarCSV 
            Caption         =   "Desde CSV"
         End
      End
      Begin VB.Menu MnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "A&yuda"
      Begin VB.Menu MnuIndice 
         Caption         =   "Indice"
      End
      Begin VB.Menu MnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAcercaDe 
         Caption         =   "Acerca de"
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
  CnnAcces.CursorLocation = adUseClient
  CnnPrincipal.CursorLocation = adUseClient
  
  Set rstUniversal = New ADODB.Recordset
  rstUniversal.CursorLocation = adUseClient
  Set rstListaPrecios = New ADODB.Recordset
  rstListaPrecios.CursorLocation = adUseClient
  
  If GetSetting("Kit Logistics", "ListasPrecios", "ConectarAlIniciar") = 1 Then
    MnuConectarAlIniciar.Checked = True
    MnuConectarBaseDatos_Click
  Else
    MnuConectarAlIniciar.Checked = False
  End If
  TxtMensaje.Width = Picture1.Width - 70
End Sub


Private Sub MnuAbrir_Click()
On Error GoTo NoValido
CDExa.Filter = "Base de datos | *.mdb|Archivos de lista de precios |*.alp"
CDExa.DialogTitle = "Lista de presios (Base de datos)"
CDExa.ShowOpen
  If CpExisteFichero(CDExa.FileName) = True Then
    RutaBaseDatos = CDExa.FileName
    CnnAcces.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & RutaBaseDatos & ";Persist Security Info=False"
    AbrirRecorset rstListaPrecios, "SELECT listaspreciosciudades.*, ciudades.NmCiudad, productos.NmProducto FROM (listaspreciosciudades LEFT JOIN ciudades ON listaspreciosciudades.IdCiudad = ciudades.IdCiudad) LEFT JOIN productos ON listaspreciosciudades.IdProducto = productos.IdProducto", CnnAcces, adOpenDynamic, adLockOptimistic
    MnuArchivo.Enabled = False
    II = 2
    FrmListas.Show
  End If
NoValido:
    If Err.Number = -2147217865 Then
      MsgBox "Esta base de datos no es un archivo valido de Precios"
      CerrarRecorset rstUniversal
      CnnPrincipal.Close
    End If
End Sub

Private Sub MnuAbrirPreciosBD_Click()
  If CpPermisoEspecial(6, CodUsuarioActivo, CnnPrincipal) = True Then
    FrmListasPrecios.Show 1
  Else
    MsgBox "No tiene permisos para ver las listas de precios", vbCritical
  End If
End Sub

Private Sub MnuAcercaDe_Click()
  FrmAcercaDe.Show 1
End Sub

Private Sub MnuConectarAlIniciar_Click()
  If MnuConectarAlIniciar.Checked = True Then
    MnuConectarAlIniciar.Checked = False
    SaveSetting "Kit logistics", "ListasPrecios", "ConectarAlIniciar", 2
  Else
    MnuConectarAlIniciar.Checked = True
    SaveSetting "Kit logistics", "ListasPrecios", "ConectarAlIniciar", 1
  End If
End Sub

Private Sub MnuConectarBaseDatos_Click()
  On Error GoTo SinConexion
    If CnnPrincipal.State = 0 Then
      CnnPrincipal.Open "DRIVER=" & GetSetting("Kit Logistics", "Configuracion", "CnnDriver") & "; SERVER=" & GetSetting("Kit Logistics", "Configuracion", "CnnServidor") & "; PORT=" & GetSetting("Kit Logistics", "Configuracion", "CnnPuerto") & "; DATABASE=" & GetSetting("Kit Logistics", "Configuracion", "CnnBaseDatos") & "; PWD=" & GetSetting("Kit Logistics", "Configuracion", "CnnContraseña") & "; UID=" & GetSetting("Kit Logistics", "Configuracion", "CnnUsuario") & ";OPTION=3"
      TxtMensaje.Text = "Conectado correctamente a BD: " & GetSetting("Kit Logistics", "Configuracion", "CnnBaseDatos") & " en Servidor: " & GetSetting("Kit Logistics", "Configuracion", "CnnServidor") & " Usuario: " & GetSetting("Kit Logistics", "Configuracion", "CnnUsuario")
      CodUsuarioActivo = IngresoSistema(CnnPrincipal, 1)
      If CodUsuarioActivo <> 0 Then
        MnuConectarBaseDatos.Enabled = False
        MnuAbrirPreciosBD.Enabled = True
      End If
  Else
    TxtMensaje.Text = "Conectado correctamente a BD: " & GetSetting("Kit Logistics", "Configuracion", "CnnBaseDatos") & " en Servidor:" & GetSetting("Kit Logistics", "Configuracion", "CnnServidor") & " Usuario:" & GetSetting("Kit Logistics", "Configuracion", "CnnUsuario")
    MnuConectarBaseDatos.Enabled = False
  End If
SinConexion:
  If Err.Number <> 0 Then
    MsgBox "No se ha podido conectar correctamente la base de datos, este error puede ser a causa de la conexion; puede hacer lo siguiente:" & Chr(13) & "- Consulte al proveedor" & Chr(13) & "- Configure la conexion desde el menu herramientas y configurar la conexion con la BD" & Chr(13) & "- Trabaje sin los modulos que requieren conexion", vbCritical, "Error de conexion"
    TxtMensaje.Text = "Error en la conexion a BD: " & GetSetting("Kit Logistics", "Configuracion", "CnnBaseDatos") & " en Servidor:" & GetSetting("Kit Logistics", "Configuracion", "CnnServidor") & " Usuario:" & GetSetting("Kit Logistics", "Configuracion", "CnnUsuario")
  End If
  
End Sub

Private Sub MnuImportarCSV_Click()
  FrmImportar.Show 1
End Sub

Private Sub MnuSalir_Click()
  Unload Me
End Sub
