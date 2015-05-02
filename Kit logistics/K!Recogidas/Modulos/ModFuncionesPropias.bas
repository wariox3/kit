Attribute VB_Name = "ModFuncionesPropias"
Option Explicit
Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, dwData As Any) As Long
Public Const HH_DISPLAY_TOPIC = &H0
Public Const HH_HELP_CONTEXT = &HF
Public Function SacarDatoString(ID As Byte) As String  'Modulok modulo para sacar el dato
  Dim i As Byte
  Open ArchivoInf For Input As #1
    For i = 0 To ID
      Input #1, FufuSt, FufuSt, SacarDatoString
    Next
Close #1
End Function
Public Function SacarConsecutivo(Campo As String) As Long
  AbrirRecorset rstUniversal, "Select " & Campo & " from Consecutivos", CnnPrincipal, adOpenDynamic, adLockOptimistic
  SacarConsecutivo = rstUniversal.Fields(Campo)
  rstUniversal.Fields(Campo) = SacarConsecutivo + 1
  rstUniversal.MoveFirst
  CerrarRecorset rstUniversal
End Function
Sub IniProg(Fin)
On Error GoTo MiError
  Principal.PgsPrincipal.Width = Principal.PicMensajes.Width
  Principal.PgsPrincipal.Min = 1
  Principal.PgsPrincipal.Max = Fin
MiError:
End Sub
Sub Prog(Valor)
On Error GoTo MiError
  Principal.PgsPrincipal.Value = Valor
MiError:
End Sub
Sub FinProg()
  Principal.PgsPrincipal.Width = 0
End Sub
Sub MsgTit(Mensaje As String)
  Principal.LblMensaje = Mensaje
  Principal.TmPrincipal.Enabled = True
End Sub
Public Function DevNombreDatosBasicos(ID As String) As String
  AbrirRecorset rstUniversal, "SELECT ID, Nombre, Apellido1, Apellido2 From DatosBasicos Where ID ='" & ID & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = False Then
    DevNombreDatosBasicos = rstUniversal.Fields("Nombre") & " " & rstUniversal.Fields("Apellido1") & " " & rstUniversal.Fields("Apellido2")
  End If
  CerrarRecorset rstUniversal
End Function

Sub ConsultaBasicos(Campo1 As String, Campo2 As String, Tabla As String)
  FufuSt = Campo1 & "," & Campo2 & "," & Tabla
  'FrmConsultaGeneral.Show 1
End Sub

Function DevEstadoDespacho(Estado As String) As String
  Select Case Estado
    Case "D"
      DevEstadoDespacho = "DIGITADO"
    Case "I"
      DevEstadoDespacho = "IMPRESO"
    Case "A"
      DevEstadoDespacho = "ANULADO"
    Case "G"
      DevEstadoDespacho = "DESCARGADO"
    Case "P"
      DevEstadoDespacho = "PROGRAMADO"
    Case "C"
      DevEstadoDespacho = "CANCELADO"
  End Select
End Function

Public Sub ResumirAsignacion(IdAsignacion As Long)
  Dim Pendientes As Long
  Dim rstAnuncios As ADODB.Recordset
  Set rstAnuncios = New ADODB.Recordset
  rstAnuncios.CursorLocation = adUseClient
  AbrirRecorset rstAnuncios, "Select count(IdAnuncio) as NroPendientes from anuncios where Efectiva=0 and IdAsignacion=" & IdAsignacion, CnnPrincipal, adOpenDynamic, adLockOptimistic
  Pendientes = rstAnuncios.Fields("NroPendientes")
  AbrirRecorset rstAnuncios, "Select count(IdAnuncio) as NroRecogidas, sum(Unidades) as Unidades, sum(KilosReales) as KilosReales, sum(KilosVol) as KilosVol from anuncios where IdAsignacion=" & IdAsignacion, CnnPrincipal, adOpenDynamic, adLockOptimistic
    AbrirRecorset rstUniversal, "Update vehiculosrecogida set Rec=" & rstAnuncios.Fields("NroRecogidas") & ", Unidades=" & Val(rstAnuncios.Fields("Unidades") & "") & ", KilosReales=" & Val(rstAnuncios.Fields("KilosReales") & "") & ", KilosVol=" & Val(rstAnuncios.Fields("KilosVol") & "") & ", Pend=" & Pendientes & " where IdAsignacion=" & IdAsignacion, CnnPrincipal, adOpenDynamic, adLockOptimistic
  CerrarRecorset rstAnuncios
End Sub
