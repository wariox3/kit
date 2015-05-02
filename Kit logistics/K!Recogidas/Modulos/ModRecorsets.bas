Attribute VB_Name = "ModRecorsets"
Option Explicit
Public Function BuscaRegistro(Criterio As String, Recorset As ADODB.Recordset) As Boolean
  Recorset.Find Criterio, , , 1
  If Recorset.EOF = True Then
    Recorset.MoveFirst
    BuscaRegistro = False
    MsgBox "No se encontro el registro"
  Else
    BuscaRegistro = True
  End If
End Function

Public Function LlenarCombo(Combo As DataCombo, ListField As String, Tabla As String)
  AbrirRecorset rstUniversal, "select " & ListField & " from " & Tabla & " ORDER BY " & ListField, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Set Combo.RowSource = rstUniversal
  Combo.ListField = ListField
End Function

Public Function ExRecorset(Fuente As String) As Boolean
  AbrirRecorset rstUniversal, Fuente, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.RecordCount > 0 Then
    ExRecorset = True
  Else
    ExRecorset = False
  End If
End Function
Public Function DevResBus(Fuente As String, Campo As String) As String
  Dim rstTem As New ADODB.Recordset
  rstTem.CursorLocation = adUseClient
  AbrirRecorset rstTem, Fuente, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstTem.EOF = False Then
    DevResBus = rstTem.Fields(Campo) & ""
  End If
  Set rstTem = Nothing
  'CerrarRecorset rstUniversal
End Function


