VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmConfiguracionGeneral 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuracion..."
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4800
      TabIndex        =   29
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton CmdAyuda 
      Caption         =   "Ayuda"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Top             =   4680
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTabConfiguracion 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "FrmConfiguracionGeneral.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Rutas"
      TabPicture(1)   =   "FrmConfiguracionGeneral.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TxtRutaImagenes"
      Tab(1).Control(1)=   "CmdCargarRutaImagenes"
      Tab(1).Control(2)=   "CmdCargarRutaReportes"
      Tab(1).Control(3)=   "TxtRutaReportes"
      Tab(1).Control(4)=   "CmdCargar"
      Tab(1).Control(5)=   "TxtArchivoAyuda"
      Tab(1).Control(6)=   "Label15"
      Tab(1).Control(7)=   "Label14"
      Tab(1).Control(8)=   "Label2"
      Tab(1).Control(9)=   "Label13"
      Tab(1).Control(10)=   "Label1"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Informacion empresa"
      TabPicture(2)   =   "FrmConfiguracionGeneral.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtResMinisterio"
      Tab(2).Control(1)=   "TxtNitEmpresa"
      Tab(2).Control(2)=   "TxtNombreEmpresa"
      Tab(2).Control(3)=   "TxtTelefonoEmpresa"
      Tab(2).Control(4)=   "TxtDireccionEmpresa"
      Tab(2).Control(5)=   "TxtNroPoliza"
      Tab(2).Control(6)=   "TxtVencePoliza"
      Tab(2).Control(7)=   "TxtNitAseguradora"
      Tab(2).Control(8)=   "TxtAseguradora"
      Tab(2).Control(9)=   "TxtDireccionTerritorial"
      Tab(2).Control(10)=   "TxtEmail"
      Tab(2).Control(11)=   "Label18"
      Tab(2).Control(12)=   "Label3"
      Tab(2).Control(13)=   "Label4"
      Tab(2).Control(14)=   "Label5"
      Tab(2).Control(15)=   "Label6"
      Tab(2).Control(16)=   "Label7"
      Tab(2).Control(17)=   "Label8"
      Tab(2).Control(18)=   "Label9"
      Tab(2).Control(19)=   "Label10"
      Tab(2).Control(20)=   "Label11"
      Tab(2).Control(21)=   "Label12"
      Tab(2).ControlCount=   22
      Begin VB.Frame Frame2 
         Caption         =   "Guias"
         Height          =   855
         Left            =   360
         TabIndex        =   44
         Top             =   2640
         Width           =   5415
         Begin VB.TextBox TxtAfectarAntesDe 
            Height          =   285
            Left            =   4200
            TabIndex        =   48
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox TxtHorasRetrasar 
            Height          =   285
            Left            =   2040
            TabIndex        =   46
            Top             =   360
            Width           =   495
         End
         Begin VB.CheckBox ChkRetrasarHoraGuias 
            Caption         =   "Retrasar hora guias"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Horas antes de las"
            Height          =   195
            Left            =   2760
            TabIndex        =   47
            Top             =   360
            Width           =   1320
         End
      End
      Begin VB.TextBox TxtResMinisterio 
         Height          =   285
         Left            =   -70680
         TabIndex        =   43
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Correo"
         Height          =   1815
         Left            =   360
         TabIndex        =   35
         Top             =   600
         Width           =   7575
         Begin VB.TextBox TxtPuertoCorreo 
            Height          =   285
            Left            =   1200
            TabIndex        =   41
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox TxtServidorCorreo 
            Height          =   285
            Left            =   1200
            TabIndex        =   40
            Top             =   480
            Width           =   6255
         End
         Begin VB.CheckBox ChkUsaSSL 
            Caption         =   "Usa SSL"
            Height          =   255
            Left            =   1200
            TabIndex        =   39
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CheckBox ChkRequiereAutenticacion 
            Caption         =   "Requiere autenticacion"
            Height          =   255
            Left            =   1200
            TabIndex        =   38
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Puerto:"
            Height          =   195
            Left            =   360
            TabIndex        =   37
            Top             =   720
            Width           =   510
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Servidor:"
            Height          =   195
            Left            =   360
            TabIndex        =   36
            Top             =   480
            Width           =   630
         End
      End
      Begin VB.TextBox TxtRutaImagenes 
         Height          =   285
         Left            =   -73560
         TabIndex        =   32
         Top             =   1680
         Width           =   5895
      End
      Begin VB.CommandButton CmdCargarRutaImagenes 
         Caption         =   "..."
         Height          =   255
         Left            =   -67560
         TabIndex        =   31
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton CmdCargarRutaReportes 
         Caption         =   "..."
         Height          =   255
         Left            =   -67560
         TabIndex        =   28
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox TxtRutaReportes 
         Height          =   285
         Left            =   -73560
         TabIndex        =   27
         Top             =   960
         Width           =   5895
      End
      Begin VB.CommandButton CmdCargar 
         Caption         =   "..."
         Height          =   255
         Left            =   -67560
         TabIndex        =   24
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox TxtArchivoAyuda 
         Height          =   285
         Left            =   -73560
         TabIndex        =   23
         Top             =   600
         Width           =   5895
      End
      Begin VB.TextBox TxtNitEmpresa 
         Height          =   285
         Left            =   -73230
         TabIndex        =   12
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox TxtNombreEmpresa 
         Height          =   285
         Left            =   -73230
         TabIndex        =   11
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox TxtTelefonoEmpresa 
         Height          =   285
         Left            =   -73230
         TabIndex        =   10
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox TxtDireccionEmpresa 
         Height          =   285
         Left            =   -73230
         TabIndex        =   9
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox TxtNroPoliza 
         Height          =   285
         Left            =   -73230
         TabIndex        =   8
         Top             =   2040
         Width           =   3735
      End
      Begin VB.TextBox TxtVencePoliza 
         Height          =   285
         Left            =   -73230
         TabIndex        =   7
         Top             =   2400
         Width           =   3735
      End
      Begin VB.TextBox TxtNitAseguradora 
         Height          =   285
         Left            =   -73230
         TabIndex        =   6
         Top             =   2760
         Width           =   3735
      End
      Begin VB.TextBox TxtAseguradora 
         Height          =   285
         Left            =   -73230
         TabIndex        =   5
         Top             =   3120
         Width           =   3735
      End
      Begin VB.TextBox TxtDireccionTerritorial 
         Height          =   285
         Left            =   -73230
         TabIndex        =   4
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox TxtEmail 
         Height          =   285
         Left            =   -73230
         TabIndex        =   3
         Top             =   3840
         Width           =   3735
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Resolucion ministerio:"
         Height          =   195
         Left            =   -72240
         TabIndex        =   42
         Top             =   3480
         Width           =   1530
      End
      Begin VB.Label Label15 
         Caption         =   "Ruta de los directorios de imagenes de conductores y vehiculos"
         Height          =   255
         Left            =   -73560
         TabIndex        =   34
         Top             =   2040
         Width           =   5535
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Ruta imagenes:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   33
         Top             =   1680
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ruta de los reportes"
         Height          =   195
         Left            =   -73560
         TabIndex        =   30
         Top             =   1320
         Width           =   1425
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Ruta reportes:"
         Height          =   195
         Left            =   -74640
         TabIndex        =   26
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Archivo de ayuda:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   25
         Top             =   600
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nit:"
         Height          =   195
         Left            =   -73770
         TabIndex        =   22
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   -74130
         TabIndex        =   21
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
         Height          =   195
         Left            =   -74250
         TabIndex        =   20
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Telefono:"
         Height          =   195
         Left            =   -74205
         TabIndex        =   19
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Numero poliza:"
         Height          =   195
         Left            =   -74580
         TabIndex        =   18
         Top             =   2040
         Width           =   1050
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Vence poliza:"
         Height          =   195
         Left            =   -74490
         TabIndex        =   17
         Top             =   2400
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nit aseguradora:"
         Height          =   195
         Left            =   -74700
         TabIndex        =   16
         Top             =   2760
         Width           =   1170
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Aseguradora:"
         Height          =   195
         Left            =   -74475
         TabIndex        =   15
         Top             =   3120
         Width           =   945
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Direccion territorial:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   14
         Top             =   3480
         Width           =   1350
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Email:"
         Height          =   195
         Left            =   -73950
         TabIndex        =   13
         Top             =   3840
         Width           =   420
      End
   End
End
Attribute VB_Name = "FrmConfiguracionGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
  Dim strInformacionEmpresa As String
  Dim strConfiguracion As String
  SaveSetting "Kit logistics", "Configuracion", "ArchivoAyuda", TxtArchivoAyuda.Text
  
  strInformacionEmpresa = "UPDATE informacionempresa SET " & _
    " Nit='" & TxtNitEmpresa.Text & "'," & _
    " Nombre='" & TxtNombreEmpresa.Text & "'," & _
    " Direccion='" & TxtDireccionEmpresa.Text & "'," & _
    " Telefono='" & TxtTelefonoEmpresa.Text & "'," & _
    " NroPoliza='" & TxtNroPoliza.Text & "'," & _
    " VencePoliza='" & TxtVencePoliza.Text & "'," & _
    " NitAseguradora='" & TxtNitAseguradora.Text & "'," & _
    " Aseguradora='" & TxtAseguradora.Text & "'," & _
    " DireccionTerritorial='" & TxtDireccionTerritorial.Text & "'," & _
    " Email='" & TxtEmail.Text & "'," & _
    " ResolucionMinTransporte='" & TxtResMinisterio.Text & "'" & _
    " WHERE Id = 1"
  AbrirRecorset rstUniversal, strInformacionEmpresa, CnnPrincipal, adOpenDynamic, adLockOptimistic
  strConfiguracion = "UPDATE configuracion SET " & _
    " ServidorCorreo='" & TxtServidorCorreo.Text & "'," & _
    " Puerto='" & TxtPuertoCorreo.Text & "', "
    If ChkRequiereAutenticacion.Value = 1 Then
      strConfiguracion = strConfiguracion & " UsaAutenticacion = 1,"
    Else
      strConfiguracion = strConfiguracion & " UsaAutenticacion = 0,"
    End If
    
    If ChkUsaSSL.Value = 1 Then
      strConfiguracion = strConfiguracion & " UsaSSL = 1,"
    Else
      strConfiguracion = strConfiguracion & " UsaSSl = 0,"
    End If
    
    If ChkRetrasarHoraGuias.Value = 1 Then
      strConfiguracion = strConfiguracion & " FechaAfectada = 1,"
    Else
      strConfiguracion = strConfiguracion & " FechaAfectada = 0,"
    End If
    strConfiguracion = strConfiguracion & " HorasAfectacion = " & Val(TxtHorasRetrasar) & ", AfectarAntesDe = " & Val(TxtAfectarAntesDe.Text)
    
  AbrirRecorset rstUniversal, strConfiguracion & " WHERE Codigo = 1", CnnPrincipal, adOpenDynamic, adLockOptimistic
  
  SaveSetting "Kit logistics", "Configuracion", "RutaReportes", TxtRutaReportes.Text
  SaveSetting "Kit logistics", "Configuracion", "DirImagenes", TxtRutaImagenes.Text
  Unload Me
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub CmdCargar_Click()
  Principal.CDExa.Filter = "Archivos de ayuda"
  Principal.CDExa.DialogTitle = "Abrir archivo (Archivo de ayuda)"
  Principal.CDExa.ShowOpen
  If Principal.CDExa.FileName <> "" Then
    TxtArchivoAyuda = Principal.CDExa.FileName
  End If
End Sub

Private Sub CmdCargarRutaImagenes_Click()
  FrmDevDirectorio.Show 1
  TxtRutaImagenes.Text = FufuSt
End Sub

Private Sub CmdCargarRutaReportes_Click()
  FrmDevDirectorio.Show 1
  TxtRutaReportes.Text = FufuSt
End Sub

Private Sub Form_Load()
  TxtArchivoAyuda.Text = GetSetting("Kit Logistics", "Configuracion", "ArchivoAyuda")
  AbrirRecorset rstUniversal, "Select informacionempresa.* from informacionempresa", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    TxtNitEmpresa.Text = rstUniversal.Fields("Nit") & ""
    TxtNombreEmpresa.Text = rstUniversal.Fields("Nombre") & ""
    TxtDireccionEmpresa.Text = rstUniversal.Fields("Direccion") & ""
    TxtTelefonoEmpresa.Text = rstUniversal.Fields("Telefono") & ""
    TxtNroPoliza.Text = rstUniversal.Fields("NroPoliza") & ""
    TxtVencePoliza.Text = Format(rstUniversal.Fields("VencePoliza") & "", "yyyy/mm/dd")
    TxtNitAseguradora.Text = rstUniversal.Fields("NitAseguradora") & ""
    TxtAseguradora.Text = rstUniversal.Fields("Aseguradora") & ""
    TxtDireccionTerritorial.Text = rstUniversal.Fields("DireccionTerritorial") & ""
    TxtEmail.Text = rstUniversal.Fields("Email") & ""
    TxtResMinisterio.Text = rstUniversal.Fields("ResolucionMinTransporte") & ""
  CerrarRecorset rstUniversal
  
  AbrirRecorset rstUniversal, "Select configuracion.* from configuracion WHERE Codigo = 1", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    TxtServidorCorreo.Text = rstUniversal.Fields("ServidorCorreo") & ""
    TxtPuertoCorreo.Text = rstUniversal.Fields("Puerto") & ""
    If Val(rstUniversal.Fields("UsaAutenticacion")) = 1 Then
      ChkRequiereAutenticacion.Value = 1
    Else
      ChkRequiereAutenticacion.Value = 0
    End If
    
    If Val(rstUniversal.Fields("UsaSSL")) = 1 Then
      ChkUsaSSL.Value = 1
    Else
      ChkUsaSSL.Value = 0
    End If
    
    If Val(rstUniversal.Fields("FechaAfectada")) = 1 Then
      ChkRetrasarHoraGuias.Value = 1
    Else
      ChkRetrasarHoraGuias.Value = 0
    End If
    TxtHorasRetrasar.Text = rstUniversal.Fields("HorasAfectacion") & ""
    TxtAfectarAntesDe.Text = rstUniversal.Fields("AfectarAntesDe") & ""
  CerrarRecorset rstUniversal
  TxtRutaReportes.Text = GetSetting("Kit Logistics", "Configuracion", "RutaReportes")
  TxtRutaImagenes.Text = GetSetting("Kit Logistics", "Configuracion", "DirImagenes")
End Sub

