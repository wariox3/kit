Attribute VB_Name = "Variables"
Option Explicit
Public CnnSeguridad As ADODB.Connection
Public rstFunPro As New ADODB.Recordset
Public TituloInf As String
Public RutaInf As String
Public TituloVentana As String
Public OpcReporte As Byte
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public SgLo As Long
Public ModIngreso As Byte

