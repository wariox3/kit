Attribute VB_Name = "ModRecorsets"
Option Explicit
Public Function ExRecorset(Fuente As String) As Boolean
  AbrirRecorset rstUniversal, Fuente, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.RecordCount > 0 Then
    ExRecorset = True
  Else
    ExRecorset = False
  End If
End Function

Public Function DevFechaHoraServidor() As Date
  Dim rstFechaHora As New ADODB.Recordset
  rstFechaHora.CursorLocation = adUseClient
  AbrirRecorset rstFechaHora, "SELECT now() as Fecha", CnnPrincipal, adOpenDynamic, adLockOptimistic
  DevFechaHoraServidor = rstFechaHora!Fecha
  Set rstFechaHora = Nothing
End Function

Public Function DevTipoCobro(TpGuia As Integer) As Integer
  AbrirRecorset rstUniversal, "SELECT TipoCobro FROM guias_tipos WHERE IdGuiaTipo = " & TpGuia, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.RecordCount > 0 Then
    DevTipoCobro = rstUniversal!TipoCobro
  Else
    DevTipoCobro = 3
  End If
End Function

Public Function DevTpGuiaFactura(TpGuia As Integer) As Integer
  AbrirRecorset rstUniversal, "SELECT GuiaFactura FROM guias_tipos WHERE IdGuiaTipo = " & TpGuia, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.RecordCount > 0 Then
    DevTpGuiaFactura = rstUniversal!GuiaFactura
  Else
    DevTpGuiaFactura = 0
  End If
End Function

Public Sub InsertarLog(IdAccion As Integer, Guia As Double)
  Dim rstInsertar As New ADODB.Recordset
  rstInsertar.CursorLocation = adUseClient
  AbrirRecorset rstInsertar, "INSERT INTO log_guias(Fecha, Guia, IdAccionLog, IdUsuario) VALUES('" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', " & Guia & ", " & IdAccion & ", " & CodUsuarioActivo & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
End Sub
