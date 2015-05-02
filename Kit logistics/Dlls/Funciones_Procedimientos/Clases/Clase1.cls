VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Recorset"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub AbrirRecorset(Recorset As ADODB.Recordset, Fuente As String, Cnn As ADODB.Connection, Cursor As CursorTypeEnum, Bloqueo As LockTypeEnum)
  On Error GoTo EnFuente
    If Recorset.State = adStateOpen Then Recorset.Close
    Recorset.Open Fuente, Cnn, Cursor, Bloqueo
EnFuente:
  If Err.Number <> 0 Then
    MsgBox Err.Description & " Error en la fuente " & Fuente
  End If
End Sub
Sub UPrimero(Recorset As ADODB.Recordset)
On Error GoTo SinReg
  Recorset.MoveFirst
SinReg:
End Sub
Sub UAnterior(Recorset As ADODB.Recordset)
  With Recorset
    If .EOF = True Then .MoveFirst
      .MovePrevious
    If .BOF = True Then .MoveFirst
  End With
End Sub
Sub USiguiente(Recorset As ADODB.Recordset)
  With Recorset
    If .BOF = True Then .MoveLast
      .MoveNext
    If .EOF = True Then .MoveLast
  End With
End Sub
Sub UUltimo(Recorset As ADODB.Recordset)
On Error GoTo SinReg
    Recorset.MoveLast
SinReg:
End Sub
Sub CerrarRecorset(Recorset As ADODB.Recordset)
  On Error GoTo ELEOF
  If Recorset.State = adStateOpen Then
    If Recorset.EOF = True Or Recorset.BOF = True Then Exit Sub
    If Recorset.EditMode = adEditAdd Or adEditInProgress Then Recorset.CancelUpdate
    Set Recorset.DataSource = Nothing
    Recorset.Close
  End If
ELEOF:
End Sub
Function SacarConsecutivo(Campo As String, Cnn As ADODB.Connection) As Long
  rstFunPro.CursorLocation = adUseClient
  AbrirRecorset rstFunPro, "Select " & Campo & " from Consecutivos", Cnn, adOpenDynamic, adLockOptimistic
  SacarConsecutivo = rstFunPro.Fields(Campo)
  rstFunPro.Fields(Campo) = SacarConsecutivo + 1
  rstFunPro.MoveFirst
  CerrarRecorset rstFunPro
End Function

Public Function DevResBus(Fuente As String, Campo As String, Cnn As ADODB.Connection) As String
  Dim rstTem As New ADODB.Recordset
  rstTem.CursorLocation = adUseClient
  AbrirRecorset rstTem, Fuente, Cnn, adOpenForwardOnly, adLockReadOnly
  If rstTem.EOF = False Then
    DevResBus = rstTem.Fields(Campo) & ""
  End If
  Set rstTem = Nothing
End Function
