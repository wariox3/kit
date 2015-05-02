Attribute VB_Name = "ModRecorsets"
Option Explicit
Public CnnPrincipal As New ADODB.Connection
Public CnnAcces As New ADODB.Connection
Public rstListaPrecios As ADODB.Recordset
Public rstUniversal As ADODB.Recordset


Public Function ExRecorset(Fuente As String) As Boolean
  AbrirRecorset rstUniversal, Fuente, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.RecordCount > 0 Then
    ExRecorset = True
  Else
    ExRecorset = False
  End If
End Function
