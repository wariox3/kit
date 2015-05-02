VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmConfiguracion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuracion"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   4440
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Liquidacion Manifiestos"
      TabPicture(0)   =   "FrmConfiguracion.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "LblTitulos(0)"
      Tab(0).Control(1)=   "LblTitulos(1)"
      Tab(0).Control(2)=   "LblTitulos(2)"
      Tab(0).Control(3)=   "TxtRteFte"
      Tab(0).Control(4)=   "TxtVrMayor"
      Tab(0).Control(5)=   "TxtIndustriaComercio"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Rutas"
      TabPicture(1)   =   "FrmConfiguracion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LblCtroOpera"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(4)=   "Label6"
      Tab(1).Control(5)=   "TxtNmCOperaciones"
      Tab(1).Control(6)=   "TxtCOperaciones"
      Tab(1).Control(7)=   "TxtRutaCoordenadasImpresionGuia"
      Tab(1).Control(8)=   "TxtRutaCoordenadasImpresionManifiesto"
      Tab(1).Control(9)=   "TxtRutaCoordenadasImpresionRecibo"
      Tab(1).Control(10)=   "TxtRutaCoordenadasImpresionPlanilla"
      Tab(1).Control(11)=   "CmdCargarRutaGuia"
      Tab(1).Control(12)=   "CmdCargarRutaManifiesto"
      Tab(1).Control(13)=   "CmdCargarRutaRecibo"
      Tab(1).Control(14)=   "CmdCargarRutaPlanilla"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Consecutivos"
      TabPicture(2)   =   "FrmConfiguracion.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(1)=   "TxtManifiestos"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "General"
      TabPicture(3)   =   "FrmConfiguracion.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "ChkImprimirGuiaFormato"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Servidor eje"
      TabPicture(4)   =   "FrmConfiguracion.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label2(0)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label7"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label8"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label2(1)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label2(2)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Label2(3)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "TxtBaseDatos"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "TxtPuerto"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "TxtDriver"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "TxtClave"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "TxtUsuario"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "TxtServidor"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).ControlCount=   12
      Begin VB.TextBox TxtServidor 
         Height          =   285
         Left            =   -73680
         TabIndex        =   32
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TxtUsuario 
         Height          =   285
         Left            =   -73680
         TabIndex        =   31
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox TxtClave 
         Height          =   285
         Left            =   -69480
         TabIndex        =   30
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox TxtDriver 
         Height          =   285
         Left            =   -73680
         TabIndex        =   29
         Top             =   1440
         Width           =   4455
      End
      Begin VB.TextBox TxtPuerto 
         Height          =   285
         Left            =   -71520
         TabIndex        =   28
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox TxtBaseDatos 
         Height          =   285
         Left            =   -69480
         TabIndex        =   27
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox ChkImprimirGuiaFormato 
         Caption         =   "Imprimir guia en formato"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton CmdCargarRutaPlanilla 
         Caption         =   "..."
         Height          =   255
         Left            =   -68040
         TabIndex        =   21
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton CmdCargarRutaRecibo 
         Caption         =   "..."
         Height          =   255
         Left            =   -68040
         TabIndex        =   20
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton CmdCargarRutaManifiesto 
         Caption         =   "..."
         Height          =   255
         Left            =   -68040
         TabIndex        =   19
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton CmdCargarRutaGuia 
         Caption         =   "..."
         Height          =   255
         Left            =   -68040
         TabIndex        =   18
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox TxtRutaCoordenadasImpresionPlanilla 
         Height          =   285
         Left            =   -72480
         TabIndex        =   17
         Top             =   2040
         Width           =   4335
      End
      Begin VB.TextBox TxtRutaCoordenadasImpresionRecibo 
         Height          =   285
         Left            =   -72480
         TabIndex        =   16
         Top             =   1680
         Width           =   4335
      End
      Begin VB.TextBox TxtRutaCoordenadasImpresionManifiesto 
         Height          =   285
         Left            =   -72480
         TabIndex        =   15
         Top             =   1320
         Width           =   4335
      End
      Begin VB.TextBox TxtRutaCoordenadasImpresionGuia 
         Height          =   285
         Left            =   -72480
         TabIndex        =   14
         Top             =   960
         Width           =   4335
      End
      Begin VB.TextBox TxtCOperaciones 
         Height          =   285
         Left            =   -73440
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TxtManifiestos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73680
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtNmCOperaciones 
         Height          =   285
         Left            =   -72480
         TabIndex        =   8
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox TxtIndustriaComercio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -72600
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TxtVrMayor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -72600
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TxtRteFte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -72600
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Driver:"
         Height          =   195
         Index           =   3
         Left            =   -74280
         TabIndex        =   38
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   2
         Left            =   -70080
         TabIndex        =   37
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Puerto:"
         Height          =   195
         Index           =   1
         Left            =   -72120
         TabIndex        =   36
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   -74400
         TabIndex        =   35
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Base datos:"
         Height          =   195
         Left            =   -70440
         TabIndex        =   34
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Servidor:"
         Height          =   195
         Index           =   0
         Left            =   -74400
         TabIndex        =   33
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ruta coordedanas planilla:"
         Height          =   195
         Left            =   -74700
         TabIndex        =   25
         Top             =   2040
         Width           =   1890
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ruta coordedanas recibo:"
         Height          =   195
         Left            =   -74655
         TabIndex        =   24
         Top             =   1680
         Width           =   1845
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ruta coordedanas manifiesto:"
         Height          =   195
         Left            =   -74925
         TabIndex        =   23
         Top             =   1320
         Width           =   2115
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ruta coordedanas guia:"
         Height          =   195
         Left            =   -74520
         TabIndex        =   22
         Top             =   960
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Manifiestos:"
         Height          =   195
         Left            =   -74640
         TabIndex        =   11
         Top             =   480
         Width           =   840
      End
      Begin VB.Label LblCtroOpera 
         AutoSize        =   -1  'True
         Caption         =   "Ctro Operaciones:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   9
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Industria y comercio:"
         Height          =   195
         Index           =   2
         Left            =   -74280
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Valor mayor para retenciones:"
         Height          =   195
         Index           =   1
         Left            =   -74880
         TabIndex        =   6
         Top             =   960
         Width           =   2100
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Retencion en la fuente:"
         Height          =   195
         Index           =   0
         Left            =   -74520
         TabIndex        =   5
         Top             =   600
         Width           =   1665
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   4440
      Width           =   1815
   End
End
Attribute VB_Name = "FrmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
  Dim strConfiguracion As String
  AbrirRecorset rstUniversal, "Update ParametrizacionLiquidaciones set RteFte=" & TxtRteFte & ", RteFteMayor=" & TxtVrMayor & ", IndCom=" & TxtIndustriaComercio, CnnPrincipal, adOpenDynamic, adLockOptimistic
  AbrirRecorset rstUniversal, "Update consecutivos set Manifiestos=" & Val(TxtManifiestos.Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
  
  SaveSetting "Kit logistics", "Movimiento", "CoordenadasImpresionGuia", TxtRutaCoordenadasImpresionGuia.Text
  SaveSetting "Kit logistics", "Movimiento", "CoordenadasImprresionManifiesto", TxtRutaCoordenadasImpresionManifiesto.Text
  SaveSetting "Kit logistics", "Movimiento", "CoordenadasImpresionReciboCaja", TxtRutaCoordenadasImpresionRecibo.Text
  SaveSetting "Kit logistics", "Movimiento", "CoordenadasImpresionPlanillaReparto", TxtRutaCoordenadasImpresionPlanilla.Text
  Coperaciones = TxtCOperaciones.Text
  SaveSetting "Kit logistics", "Configuracion", "Coperaciones", Coperaciones
  
  strConfiguracion = "UPDATE configuracion SET "
  If ChkImprimirGuiaFormato.value = 1 Then
    strConfiguracion = strConfiguracion & " ImprimirGuiaFormato = 1"
  Else
    strConfiguracion = strConfiguracion & " ImprimirGuiaFormato = 0"
  End If
  strConfiguracion = strConfiguracion & ", ejeServidor = '" & TxtServidor.Text & "', ejePuerto='" & TxtPuerto.Text & "', ejeBaseDatos='" & TxtBaseDatos.Text & "', ejeUsuario='" & TxtUsuario.Text & "', ejeClave='" & TxtClave.Text & "', ejeDriver='" & TxtDriver.Text & "' "
  AbrirRecorset rstUniversal, strConfiguracion & " WHERE Codigo = 1", CnnPrincipal, adOpenDynamic, adLockOptimistic
  
  MsgBox "cambios aplicados con exito", vbInformation
  Unload Me
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub CmdCargarRutaReportes_Click()
  
End Sub

Private Sub CmdCargarRutaGuia_Click()
    Principal.CDExa.Filter = "Archivo Coordenadas Remision |*.rmci"
    Principal.CDExa.DialogTitle = "Archivo Coordenadas (Remision)"
    Principal.CDExa.ShowOpen
    TxtRutaCoordenadasImpresionGuia.Text = Principal.CDExa.FileName
End Sub

Private Sub CmdCargarRutaManifiesto_Click()
    Principal.CDExa.Filter = "Archivo Coordenadas Manifiesto |*.mfto"
    Principal.CDExa.DialogTitle = "Archivo Coordenadas (Manifiesto)"
    Principal.CDExa.ShowOpen
    TxtRutaCoordenadasImpresionManifiesto.Text = Principal.CDExa.FileName
End Sub

Private Sub CmdCargarRutaPlanilla_Click()
    Principal.CDExa.Filter = "Archivo Coordenadas orden de recogida |*.ork"
    Principal.CDExa.DialogTitle = "Archivo Coordenadas (Orden de recogida)"
    Principal.CDExa.ShowOpen
    TxtRutaCoordenadasImpresionPlanilla.Text = Principal.CDExa.FileName
End Sub

Private Sub CmdCargarRutaRecibo_Click()
    Principal.CDExa.Filter = "Archivo Coordenadas Recibo de caja |*.rcs"
    Principal.CDExa.DialogTitle = "Archivo Coordenadas (Recibo de caja)"
    Principal.CDExa.ShowOpen
    TxtRutaCoordenadasImpresionRecibo.Text = Principal.CDExa.FileName
End Sub

Private Sub Form_Load()
  TxtCOperaciones.Text = Coperaciones
  AbrirRecorset rstUniversal, "Select IdPO, NmPuntoOperaciones from CentrosOperaciones where IdPO=" & Coperaciones, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      TxtNmCOperaciones.Text = rstUniversal.Fields("NmPuntoOperaciones")
    End If
  CerrarRecorset rstUniversal
  
  AbrirRecorset rstUniversal, "Select*from ParametrizacionLiquidaciones", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    TxtRteFte = rstUniversal!RteFte
    TxtVrMayor = rstUniversal!RteFteMayor
    TxtIndustriaComercio = rstUniversal!IndCom
  CerrarRecorset rstUniversal
  
  AbrirRecorset rstUniversal, "Select*from Consecutivos", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    TxtManifiestos.Text = rstUniversal!Manifiestos
  CerrarRecorset rstUniversal
    
  TxtRutaCoordenadasImpresionGuia.Text = GetSetting("Kit Logistics", "Movimiento", "CoordenadasImpresionGuia")
  TxtRutaCoordenadasImpresionManifiesto.Text = GetSetting("Kit Logistics", "Movimiento", "CoordenadasImprresionManifiesto")
  TxtRutaCoordenadasImpresionRecibo.Text = GetSetting("Kit Logistics", "Movimiento", "CoordenadasImpresionReciboCaja")
  TxtRutaCoordenadasImpresionPlanilla.Text = GetSetting("Kit Logistics", "Movimiento", "CoordenadasImpresionPlanillaReparto")
  
  AbrirRecorset rstUniversal, "Select configuracion.* FROM configuracion", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If Val(rstUniversal.Fields("ImprimirGuiaFormato")) = 1 Then
      ChkImprimirGuiaFormato.value = 1
    Else
      ChkImprimirGuiaFormato.value = 0
    End If
    TxtServidor.Text = rstUniversal!ejeServidor & ""
    TxtPuerto.Text = rstUniversal!ejePuerto & ""
    TxtBaseDatos.Text = rstUniversal!ejeBaseDatos & ""
    TxtUsuario.Text = rstUniversal!ejeUsuario & ""
    TxtClave.Text = rstUniversal!ejeClave & ""
    TxtDriver.Text = rstUniversal!ejeDriver & ""
  CerrarRecorset rstUniversal
  

  
End Sub

