VERSION 5.00
Begin VB.UserControl ToolConsultas 
   CanGetFocus     =   0   'False
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   465
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   435
   ScaleWidth      =   465
   ToolboxBitmap   =   "ToolEspecial.ctx":0000
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "ToolEspecial.ctx":0312
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "ToolConsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'Default Property Values:
Const m_def_Fecha1 = 0
Const m_def_Fecha2 = 0
Const m_def_DatLo = 0
Const m_def_DatSt = "0"
'Property Variables:
Dim m_Fecha1 As Date
Dim m_Fecha2 As Date
Dim m_DatLo As Long
Dim m_DatSt As String
'Event AbrirDevDatos()

Public Function AbrirDevDatos(Titulo As String, Mensaje As String, Tipo As Integer, Defecto As Long) As Boolean
  FrmDev.Caption = Titulo
  FrmDev.LblMensaje = Mensaje
  Tip = Tipo
  Load FrmDev
  Select Case Tip
    Case 1 'Contraseñas
      FrmDev.TxtIngreso.PasswordChar = "*"
      FrmDev.TxtIngreso.MaxLength = 20
    Case 2 'Doc Clientes
      FrmDev.TxtIngreso.MaxLength = 20
    Case 3 'Long
      If Defecto > 0 Then
        FrmDev.TxtIngreso.Text = Defecto
      End If
      FrmDev.TxtIngreso.MaxLength = 10
    Case 4  'Placas
      FrmDev.TxtIngreso.MaxLength = 6
    Case 5 'Nits
      FrmDev.TxtIngreso.MaxLength = 10
    Case 6 'Guia nueva logicuartas
      If Defecto > 0 Then
        FrmDev.TxtIngreso.Text = Defecto
      End If
      FrmDev.TxtIngreso.MaxLength = 20
  End Select
  FrmDev.Show 1
  DatSt = ElSt
  DatLo = Ello
  AbrirDevDatos = Ok
End Function

'Event AbrirDevFechas()
Public Function AbrirDevFechas(Titulo As String, Mensaje As String, NroFechas As Byte) As Boolean
  FrmDevFechas.Caption = Titulo
  FrmDevFechas.LblMensaje = Mensaje
  FrmDevFechas.Tag = NroFechas
  FrmDevFechas.Show 1
  Fecha1 = LasFechas(1)
  Fecha2 = LasFechas(2)
  AbrirDevFechas = Ok
End Function
Public Function AbrirDevConsulta(Tipo As Byte, Cnn As ADODB.Connection) As Boolean
  Dim TP As Byte
  Set Cn = Cnn
  TP = 0
  Select Case Tipo
    Case 1 'Consulta Ciudades
      FrmConsultaCiudades.Show 1
      TP = 2
    Case 2 'Consulta Cliente
      FrmConsultaCliente.Show 1
      TP = 2
    Case 3  'Consulta Conductores
      FrmConsultaConductores.Show 1
      TP = 1
    Case 4 'Consulta Remitentes
      FrmConsultaRemitentes.Show 1
      TP = 1
    Case 5 'Consulta Vehiculos
      FrmConsultaVehiculos.Show 1
      TP = 1
    Case 6 'Control Post
      FrmConsultaControlPost.Show 1
      TP = 2
    Case 7 ' Terceros
      FrmConsultaTerceros.Show 1
      TP = 1
    Case 8 'COnsulta DatosBasicos
      FrmConsultaDatosBasicos.Show 1
      TP = 1
    Case 9 'COnsulta Destinatarios
      FrmConsultaDestinatarios.Show 1
      TP = 1
    Case 10 'COnsulta auxiliares
      FrmConsultaAuxiliares.Show 1
      TP = 1
    Case 11 'Listas de Precios
      FrmConsultaListasPrecios.Show 1
      TP = 2
  End Select
  
  If TP = 1 Then
    If ElSt = "" Then
      AbrirDevConsulta = False
      DatSt = ""
    Else
      AbrirDevConsulta = True
      DatSt = ElSt
    End If
  Else
    If Ello = 0 Then
      AbrirDevConsulta = False
      DatLo = 0
    Else
      AbrirDevConsulta = True
      DatLo = Ello
    End If
  End If
End Function
Public Function AbrirConsultaGral(Campo1 As String, Campo2 As String, Tabla As String, Cnn As ADODB.Connection) As Boolean
  Set Cn = Cnn
  ElSt = Campo1 & "," & Campo2 & "," & Tabla
  FrmConsultaGeneral.Show 1
  If Ello <> 0 Then
    DatLo = Ello
    AbrirConsultaGral = True
  Else
    DatLo = 0
    AbrirConsultaGral = False
  End If
End Function

Public Function AbrirDevConsultaCO(Tipo As Byte, CO As Long, Cnn As ADODB.Connection) As Boolean
  Dim TP As Byte
  Set Cn = Cnn
  TP = 0
  Coperaciones = CO
  Select Case Tipo
    Case 1 'Consulta Rutas Urbanas
      FrmConsultaRutasUrbanas.Show 1
      TP = 2
    Case 2 'Consulta Rutas
      FrmConsultaRutas.Show 1
      TP = 2
  End Select
  
  If TP = 1 Then
    If ElSt = "" Then
      AbrirDevConsultaCO = False
      DatSt = ""
    Else
      AbrirDevConsultaCO = True
      DatSt = ElSt
    End If
  Else
    If Ello = 0 Then
      AbrirDevConsultaCO = False
      DatLo = 0
    Else
      AbrirDevConsultaCO = True
      DatLo = Ello
    End If
  End If
End Function




Private Sub UserControl_Resize()
  Image1.Top = 0
  Image1.Left = 0
  UserControl.Height = Image1.Height
  UserControl.Width = Image1.Width
End Sub
Public Property Get Fecha1() As Date
Attribute Fecha1.VB_Description = "Datos ""Date1"" que puede devolver la ventana de [DevFechas]"
Attribute Fecha1.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
  Fecha1 = m_Fecha1
End Property
Public Property Let Fecha1(ByVal New_Fecha1 As Date)
  m_Fecha1 = New_Fecha1
  PropertyChanged "Fecha1"
End Property
Public Property Get Fecha2() As Date
Attribute Fecha2.VB_Description = "Datos ""Date2"" que puede devolver la ventana de [DevFechas]"
Attribute Fecha2.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
  Fecha2 = m_Fecha2
End Property
Public Property Let Fecha2(ByVal New_Fecha2 As Date)
  m_Fecha2 = New_Fecha2
  PropertyChanged "Fecha2"
End Property
Public Property Get DatLo() As Long
Attribute DatLo.VB_Description = "Datos ""long"" que puede devolver la ventana de [DevDatos]"
Attribute DatLo.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
  DatLo = m_DatLo
End Property
Public Property Let DatLo(ByVal New_DatLo As Long)
  m_DatLo = New_DatLo
  PropertyChanged "DatLo"
End Property
Public Property Get DatSt() As String
Attribute DatSt.VB_Description = "Datos ""String"" que puede devolver la ventana de [DevDatos]"
Attribute DatSt.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
  DatSt = m_DatSt
End Property
Public Property Let DatSt(ByVal New_DatSt As String)
  m_DatSt = New_DatSt
  PropertyChanged "DatSt"
End Property
Private Sub UserControl_InitProperties()
  m_Fecha1 = m_def_Fecha1
  m_Fecha2 = m_def_Fecha2
  m_DatLo = m_def_DatLo
  m_DatSt = m_def_DatSt
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Fecha1 = PropBag.ReadProperty("Fecha1", m_def_Fecha1)
  m_Fecha2 = PropBag.ReadProperty("Fecha2", m_def_Fecha2)
  m_DatLo = PropBag.ReadProperty("DatLo", m_def_DatLo)
  m_DatSt = PropBag.ReadProperty("DatSt", m_def_DatSt)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Fecha1", m_Fecha1, m_def_Fecha1)
  Call PropBag.WriteProperty("Fecha2", m_Fecha2, m_def_Fecha2)
  Call PropBag.WriteProperty("DatLo", m_DatLo, m_def_DatLo)
  Call PropBag.WriteProperty("DatSt", m_DatSt, m_def_DatSt)
End Sub
