Attribute VB_Name = "ModFuncionesRecorset"
Option Explicit
Public Function ExRecorset(Fuente As String) As Boolean
  AbrirRecorset rstUniversal, Fuente, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.RecordCount > 0 Then
    ExRecorset = True
  Else
    ExRecorset = False
  End If
End Function

Public Function BuscaRegistro(Criterio As String, Recorset As ADODB.Recordset) As Boolean
  Recorset.Find Criterio, , , 1
  If Recorset.EOF = True Then
    Recorset.MoveFirst
    BuscaRegistro = False
    MsgTit "No se encontro el registro"
  Else
    BuscaRegistro = True
  End If
End Function
