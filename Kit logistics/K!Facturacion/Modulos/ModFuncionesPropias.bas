Attribute VB_Name = "ModFuncionesPropias"
Option Explicit

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
Function DevEstado(Estado As String) As String
  Select Case Estado
    Case "D"
      DevEstado = "DIGITADA"
    Case "I"
      DevEstado = "IMPRESA"
    Case "A"
      DevEstado = "ANULADA"
    Case "C"
      DevEstado = "CANCELADA"
    Case "B"
      DevEstado = "ABONADA"
  End Select
End Function
Public Function DigitoNIT(ByVal sNit As String) As String
    On Error Resume Next
    Dim sTMP, sTmp1, sTmp2, aux As String
    Dim I As Integer
    Dim iResiduo  As Integer
    Dim iChequeo As Integer
    Dim iPrimos(15) As Integer '<- Defino el Arreglo de los Primos.
    For I = 1 To Len(sNit)
      If Mid(sNit, I, 1) <> "-" Then
        aux = aux & Mid(sNit, I, 1)
      End If
    Next I
    sNit = aux
    
    iPrimos(1) = 3: iPrimos(2) = 7: iPrimos(3) = 13: iPrimos(4) = 17: iPrimos(5) = 19
    iPrimos(6) = 23: iPrimos(7) = 29: iPrimos(8) = 37: iPrimos(9) = 41: iPrimos(10) = 43
    iPrimos(11) = 47: iPrimos(12) = 53: iPrimos(13) = 59: iPrimos(14) = 67: iPrimos(15) = 71
    iChequeo = 0: iResiduo = 0
    For I = 0 To Len(Trim(sNit)) - 1
        sTMP = Mid(sNit, Len(Trim(sNit)) - I, 1)
        iChequeo = iChequeo + (Val(sTMP) * iPrimos(I + 1))
        'MsgBox Val(sTmp), vbCritical, iPrimos(i + 1)
    Next I
    iResiduo = iChequeo Mod 11
    If iResiduo <= 1 Then
        If iResiduo = 0 Then DigitoNIT = 0
        If iResiduo = 1 Then DigitoNIT = 1
    Else
        DigitoNIT = 11 - iResiduo
    End If
    DigitoNIT = aux & DigitoNIT
    'By GeNeTiKo
End Function
Public Function DigitoVerificacion(ByVal sNit As String) As String
    On Error Resume Next
    Dim sTMP, sTmp1, sTmp2, aux As String
    Dim I As Integer
    Dim iResiduo  As Integer
    Dim iChequeo As Integer
    Dim iPrimos(15) As Integer '<- Defino el Arreglo de los Primos.
    For I = 1 To Len(sNit)
      If Mid(sNit, I, 1) <> "-" Then
        aux = aux & Mid(sNit, I, 1)
      End If
    Next I
    sNit = aux
    
    iPrimos(1) = 3: iPrimos(2) = 7: iPrimos(3) = 13: iPrimos(4) = 17: iPrimos(5) = 19
    iPrimos(6) = 23: iPrimos(7) = 29: iPrimos(8) = 37: iPrimos(9) = 41: iPrimos(10) = 43
    iPrimos(11) = 47: iPrimos(12) = 53: iPrimos(13) = 59: iPrimos(14) = 67: iPrimos(15) = 71
    iChequeo = 0: iResiduo = 0
    For I = 0 To Len(Trim(sNit)) - 1
        sTMP = Mid(sNit, Len(Trim(sNit)) - I, 1)
        iChequeo = iChequeo + (Val(sTMP) * iPrimos(I + 1))
        'MsgBox Val(sTmp), vbCritical, iPrimos(i + 1)
    Next I
    iResiduo = iChequeo Mod 11
    If iResiduo <= 1 Then
        If iResiduo = 0 Then DigitoVerificacion = 0
        If iResiduo = 1 Then DigitoVerificacion = 1
    Else
        DigitoVerificacion = 11 - iResiduo
    End If
    DigitoVerificacion = DigitoVerificacion
    'By GeNeTiKo
End Function

Public Sub ExportarExcel(rstTemp As ADODB.Recordset)
  Dim RutaSalida As String
  Dim o_Excel     As Object
  Dim o_Libro     As Object
  Dim o_Hoja      As Object
  Dim Fila        As Long
  Dim Columna     As Long
  
On Error GoTo Error_Handler

  Principal.CDExa.DialogTitle = "Guardar como"
  Principal.CDExa.Filter = "Archivo Excel|*.xls"
  Principal.CDExa.ShowSave
  If Principal.CDExa.FileName <> "" Then
    RutaSalida = Principal.CDExa.FileName
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
    For Columna = 1 To rstTemp.Fields.Count
      o_Hoja.Cells(1, Columna).Value = rstTemp.Fields(Columna - 1).Name
    Next
    
    With rstTemp
        For Fila = 2 To .RecordCount + 1
            For Columna = 0 To .Fields.Count - 1
                o_Hoja.Cells(Fila, Columna + 1).Value = .Fields(Columna).Value
            Next
            .MoveNext
        Next
    End With
    o_Libro.Close True, RutaSalida
    o_Excel.Quit
    
  End If
  Exit Sub
Error_Handler:
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
        
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Sub
