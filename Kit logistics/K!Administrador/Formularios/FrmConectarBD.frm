VERSION 5.00
Begin VB.Form FrmConectarBD 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Conectar con la Base de Datos"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtPuerto 
      Height          =   285
      Left            =   1080
      TabIndex        =   13
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox TxtDriver 
      Height          =   285
      Left            =   1080
      TabIndex        =   12
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "Guardar"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton CmdConectar 
      Caption         =   "Comprobar"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox TxtContraseña 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox TxtUsuario 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox TxtBaseDeDatos 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox TxtServidor 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Puerto:"
      Height          =   195
      Left            =   480
      TabIndex        =   14
      Top             =   840
      Width           =   510
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Controlador:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.Label LblEstado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   3975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Servidor:"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Base Datos:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Usuario:"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "FrmConectarBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CnnComprobar As New ADODB.Connection

Private Sub CmdConectar_Click()
  On Error GoTo ErrCnn
    CnnComprobar.Open "DRIVER=" & TxtDriver.Text & "; SERVER=" & TxtServidor & "; PORT=" & TxtPuerto.Text & "; DATABASE=" & TxtBaseDeDatos & "; PWD=" & TxtContraseña & "; UID=" & TxtUsuario & ";OPTION=3"
ErrCnn:
    If Err.Number <> 0 Then
      MsgBox "No es posible conectar con esta cadena de conexion: " & Err.Description, vbCritical
      LblEstado.Caption = "Error al conectar"
    Else
      MsgBox "Conexion establecida con exito", vbInformation
      LblEstado.Caption = "Conexion establecida satisfactoriamente"
    End If
    Set CnnComprobar = Nothing
End Sub

Private Sub CmdGuardar_Click()
    SaveSetting "Kit logistics", "Configuracion", "CnnServidor", TxtServidor.Text
    SaveSetting "Kit logistics", "Configuracion", "CnnBaseDatos", TxtBaseDeDatos.Text
    SaveSetting "Kit logistics", "Configuracion", "CnnUsuario", TxtUsuario.Text
    SaveSetting "Kit logistics", "Configuracion", "CnnCOntraseña", TxtContraseña.Text
    SaveSetting "Kit logistics", "Configuracion", "CnnDriver", TxtDriver.Text
    SaveSetting "Kit logistics", "Configuracion", "CnnPuerto", TxtPuerto.Text
End Sub

Private Sub Form_Load()
  CnnComprobar.CursorLocation = adUseClient
  TxtServidor.Text = GetSetting("Kit Logistics", "Configuracion", "CnnServidor")
  TxtBaseDeDatos.Text = GetSetting("Kit Logistics", "Configuracion", "CnnBaseDatos")
  TxtUsuario.Text = GetSetting("Kit Logistics", "Configuracion", "CnnUsuario")
  TxtContraseña.Text = GetSetting("Kit Logistics", "Configuracion", "CnnContraseña")
  TxtDriver.Text = GetSetting("Kit Logistics", "Configuracion", "CnnDriver")
  TxtPuerto.Text = GetSetting("Kit Logistics", "Configuracion", "CnnPuerto")
  LblEstado.Caption = "Sin comproar conexion"
End Sub

