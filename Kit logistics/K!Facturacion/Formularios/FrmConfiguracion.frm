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
      TabIndex        =   1
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   4440
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   135
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Consecutivos"
      TabPicture(0)   =   "FrmConfiguracion.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TxtNotasCredito"
      Tab(0).Control(1)=   "TxtRecibos"
      Tab(0).Control(2)=   "TxtFacturas"
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(5)=   "Label2"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Rutas"
      TabPicture(1)   =   "FrmConfiguracion.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "CmdCargarRutaFactura"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "TxtRutaCoordenadasImpresionFactura"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "TxtRutaExportarFacturas"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.TextBox TxtRutaExportarFacturas 
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Top             =   960
         Width           =   4335
      End
      Begin VB.TextBox TxtRutaCoordenadasImpresionFactura 
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Top             =   600
         Width           =   4335
      End
      Begin VB.CommandButton CmdCargarRutaFactura 
         Caption         =   "..."
         Height          =   255
         Left            =   6600
         TabIndex        =   9
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox TxtNotasCredito 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73560
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox TxtRecibos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73560
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtFacturas 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73560
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ruta exportar archivo fac:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1845
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ruta coordedanas factura:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Notas Credito:"
         Height          =   195
         Left            =   -74775
         TabIndex        =   8
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Recibos:"
         Height          =   195
         Left            =   -74400
         TabIndex        =   6
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Facturas:"
         Height          =   195
         Left            =   -74430
         TabIndex        =   5
         Top             =   960
         Width           =   660
      End
   End
End
Attribute VB_Name = "FrmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
  AbrirRecorset rstUniversal, "Update consecutivos set Facturas=" & Val(TxtFacturas.Text) & ", Recibos = " & Val(TxtRecibos.Text) & ", NotasCredito = " & Val(TxtNotasCredito.Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
  SaveSetting "Kit logistics", "Facturacion", "CoordenadasImpresionFactura", TxtRutaCoordenadasImpresionFactura.Text
  SaveSetting "Kit logistics", "Facturacion", "RutaExportarArchivoFacturas", TxtRutaExportarFacturas.Text
  MsgBox "cambios aplicados con exito", vbInformation
  Unload Me
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub CmdCargarRutaFactura_Click()
    Principal.CDExa.Filter = "Archivo Coordenadas factura |*.fsl"
    Principal.CDExa.DialogTitle = "Archivo Coordenadas (factura)"
    Principal.CDExa.ShowOpen
    TxtRutaCoordenadasImpresionFactura.Text = Principal.CDExa.FileName
End Sub

Private Sub Form_Load()
  AbrirRecorset rstUniversal, "Select*from Consecutivos", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    TxtRecibos.Text = rstUniversal!Recibos
    TxtFacturas.Text = rstUniversal!Facturas
    TxtNotasCredito.Text = rstUniversal!NotasCredito
  CerrarRecorset rstUniversal
  
  TxtRutaCoordenadasImpresionFactura.Text = GetSetting("Kit Logistics", "Facturacion", "CoordenadasImpresionFactura")
  TxtRutaExportarFacturas.Text = GetSetting("Kit Logistics", "Facturacion", "RutaExportarArchivoFacturas")

End Sub

