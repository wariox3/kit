Attribute VB_Name = "ModVariables"
Option Explicit

'*******************************************
Public CnnPrincipal As ADODB.Connection
Public CnnBackut As ADODB.Connection
Public rstUniversal As ADODB.Recordset
Public rstUniversalSer As ADODB.Recordset
Public rstUniversalAux As ADODB.Recordset
'*******************************************

Public CodUsuarioActivo As Integer
Public CodEmpresaActiva As Integer
Public UsuarioActivo As String
Public EmpresaActiva As String
Public RutaLocal As String
Public ArchivoInf As String
Public Coperaciones As Long
Public MProductos(5) As EstructuraMatrizProductos
Public IdClienteViejo As String
Public CooImp() As CoordenadasImpresion
Public GuiaManConsecutivo As Boolean

Public Permisos() As TipPermiso
Public II As Long
Public Item As ListItem

Public FufuSt As String
Public FufuLo As Long
Public FufuLo2 As Long
Public FufuDo As Double
Public GuiaDesde As Double
Public GuiaHasta As Double

Type TipListaPrecios
  Devuelve As Boolean
  VrKilo As Currency
  KTope As Long
  VrKTope As Currency
  VrKdicional As Currency
  VrUnidad As Currency
  Minimos As Integer
End Type

Type EstructuraMatrizProductos
  IdProducto As Integer
  IdEmpaque As Integer
  Ancho As Long
  Largo As Long
  Alto As Long
  KilosVol As Long
  kilosReales As Long
  KilosFacturados As Long
  Cantidad As Integer
  VrFlete As Currency
  Lote As String
  NmEmpaque As String
  NmProducto As String
End Type


Public Type CoordenadasImpresion
  VX As Integer
  VY As Integer
  Mostrar As Byte
  Tamaño As Byte
  Longitud As Integer
End Type

Public Type TipPermiso
  Formulario As Integer
  Ingreso As Boolean
  Nuevo As Boolean
  Editar As Boolean
  Eliminar As Boolean
End Type

