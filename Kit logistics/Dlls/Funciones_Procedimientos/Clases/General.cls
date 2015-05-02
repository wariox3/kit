VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "General"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim Numero(18) As String
'Public Const Concatenar = "y"         'Palabra que Concatena los Números
Dim Punto As String                   'Palabra para el punto decimal
'Public Const cientos = "cientos"      'para las centenas
'Public Const millares = "mil"         'para los millares
'Public Const millones = "millones"    'para los millones
Dim Decenas(8) As String              'para el 20,30,40,50,60,70,80,90
Dim CentenasEspeciales(9) As String   '100,500,700,900
Const ENCRYPT = 1
Const DECRYPT = 2


Public Function CpExisteFichero(Ruta As String) As Boolean
  Dim X
  On Error GoTo ErrorHandler:
  X = GetAttr(Ruta)
  CpExisteFichero = True
  Exit Function
ErrorHandler:
  CpExisteFichero = False
End Function
Public Function MostrarMensaje(mm As String) As Boolean
  MsgBox mm
  MostrarMensaje = False
End Function

Public Function IngresoSistema(Cn As ADODB.Connection, Modulo As Byte) As Long
  Set CnnSeguridad = Cn
  ModIngreso = Modulo
  FrmContraseña.Show 1
  IngresoSistema = SgLo
End Function

Sub ValidarEntrada(Caja, Tecla As Integer, Tipo As Byte)
  If Tipo = 1 Then 'Solo permite ingresar numeros
    If InStr("0123456789" + Chr(8) + Chr(13), Chr(Tecla)) = 0 Then Tecla = 0: MsgBox "Solo se permiten numeros en este campo"
  ElseIf Tipo = 2 Then 'Solo permite ingresar numeros y punto
    If InStr("0123456789." + Chr(8) + Chr(13), Chr(Tecla)) = 0 Then Tecla = 0: MsgBox "Solo se permiten numero y punto en este campo"
  ElseIf Tipo = 3 Then 'Porcentaje
    '********************************
      Dim i As Byte, I1 As Byte
      If InStr("0123456789." + Chr(8) + Chr(13), Chr(Tecla)) = 0 Then Tecla = 0: MsgBox "Solo se permiten numero y punto en este campo": Exit Sub
      If Tecla = 46 And Len(Caja.Text) < 1 Then Tecla = 0
      If Caja <> "" Then
        If Val(Caja.Text & Chr(Tecla)) >= 101 Then
          Tecla = 0
          Exit Sub
        End If
    
        If Mid(Caja.Text, Len(Caja.Text), 1) = "." And Tecla = 46 Then
          Tecla = 0
          Exit Sub
        End If
        If Tecla = 46 Then
          If Val(Caja.Text) >= 100 Then
            Tecla = 0
            Exit Sub
          End If
          For I1 = 1 To Len(Caja.Text)
            If Mid(Caja.Text, I1, 1) = "." Then
              Tecla = 0
              Exit For
            End If
          Next
        End If
      End If
      If Tecla <> 8 Then
        For I1 = 1 To Len(Caja.Text)
          If Mid(Caja.Text, I1, 1) = "." Then
            If Len(Caja.Text) - I1 > 1 Then Tecla = 0
          End If
        Next
      End If
    '********************************
  ElseIf Tipo = 4 Then 'Tipo de documento
    If InStr("NCTE" + Chr(8) + Chr(13), Chr(Tecla)) = 0 Then Tecla = 0: MsgBox "Solo se permite N-Nit C-Cedula, T-Tarjeta, E-Extranjeria, en este campo [MAYUSCULAS]"
  ElseIf Tipo = 5 Then 'Numeros con negativo y positivo
      If InStr("-0123456789" + Chr(8) + Chr(13), Chr(Tecla)) = 0 Then Tecla = 0: MsgBox "Solo se permiten numeros y (-) en este campo"
  End If
End Sub
Function DevCheck(Res As Boolean) As Byte
  If Res = True Then
    DevCheck = 1
  Else
    DevCheck = 0
  End If
End Function

Public Sub LlenarNumeros()
Punto = " con "                   'Como se Leerá el Punto
Numero(0) = ""
Numero(1) = "Uno"
Numero(2) = "Dos"
Numero(3) = "Tres"
Numero(4) = "Cuatro"
Numero(5) = "Cinco"
Numero(6) = "Seis"
Numero(7) = "Siete"
Numero(8) = "Ocho"
Numero(9) = "Nueve"
Numero(10) = ""
Numero(11) = "Once"
Numero(12) = "Doce"
Numero(13) = "Trece"
Numero(14) = "Catorce"
Numero(15) = "Quince"
'Números con pronunciaciónes Especiales
Numero(16) = "Dieci"
Numero(17) = "Veinti"

'Nombres para las Decenas Básicas
Decenas(0) = "Diez"
Decenas(1) = "Veinte"
Decenas(2) = "Treinta"
Decenas(3) = "Cuarenta"
Decenas(4) = "Cincuenta"
Decenas(5) = "Sesenta"
Decenas(6) = "Setenta"
Decenas(7) = "Ochenta"
Decenas(8) = "Noventa"

'Centenas Especiales
CentenasEspeciales(0) = "Cien"
CentenasEspeciales(1) = "Ciento"
CentenasEspeciales(4) = "Quinientos"
CentenasEspeciales(6) = "Setecientos"
CentenasEspeciales(8) = "Novecientos"

End Sub

Public Function CovLetras(num As String) As String
LlenarNumeros
Dim enteros As Integer      'Longitud del numero entero
Dim decimales As Integer    'Numeros decimales
Dim p_entera As String      'Parte Entera del Número
Dim p_decimal As String     'Parte decimal

If InStr(1, num, ".") = 0 Then  'Verificamos si existe punto decimal
    enteros = CInt(Len(num))
Else
    enteros = CInt(Len(Mid(num, 1, InStr(1, num, ".") - 1)))
    decimales = CInt(Len(Trim(Mid(num, InStr(1, num, ".") + 1, Len(num)))))
End If

'Primero Traducimos la Parte Entera del Número
p_entera = LeerNumero(Mid(num, 1, enteros))

'Luego, si hay decimales, los traducimos
If decimales > 0 Then
    p_decimal = Punto & LeerNumero(Trim(Mid(num, InStr(1, num, ".") + 1, Len(num))))
End If

CovLetras = p_entera & p_decimal
End Function
Private Function LeerNumero(num As String) As String
'Funcion Recursiva que devuelve el número, traducido

On Error GoTo error_leernumero
Dim level As Integer    'decenas,centenas,  "mil" ..
Dim Texto As String     'texto numérico

level = Len(num)        'Longitud de la cadena

Select Case level       'Segun sea el Nivel
    Case 1:             'Nivel Básico
            Texto = Texto & Numero(CInt(num))
    Case 2:             'Nivel Decimal
            Select Case CInt(num)
                Case 11 To 15:
                    Texto = Texto & Numero(CInt(num))
                Case 16 To 19
                    Texto = Numero(16) & LCase(LeerNumero(Mid(num, 2, 1)))
                Case 10, 20, 30, 40, 50, 60, 70, 80, 90
                    Texto = Decenas(CInt(Mid(num, 1, 1)) - 1)
                Case 20 To 29:
                    Texto = Numero(17) & LCase(LeerNumero(Mid(num, 2, 1)))
                Case Else:
                    Texto = Decenas(CInt(Mid(num, 1, 1)) - 1) & " " & "y" & " " & LCase(LeerNumero(Mid(num, 2, 1)))
            End Select
    Case 3:             'Nivel Centenar
            Select Case CInt(num)
                Case 100, 500, 700, 900:
                    Texto = CentenasEspeciales(CInt(Mid(num, 1, 1)) - 1)
                Case 200, 300, 400, 600, 800:
                    Texto = Numero(CInt(Mid(num, 1, 1))) & "cientos"
                Case 101 To 199:
                    Texto = CentenasEspeciales(1) & " " & LeerNumero(CStr(CInt(Mid(num, 2, 2))))
                Case 501 To 599, 701 To 799, 901 To 999:
                    Texto = CentenasEspeciales(CInt(Mid(num, 1, 1)) - 1) & " " & LeerNumero(CStr(CInt(Mid(num, 2, 2))))
                Case Else:
                    Texto = Numero(CInt(Mid(num, 1, 1))) & "cientos" & " " & LeerNumero(CStr(CInt(Mid(num, 2, 2))))
            End Select
    Case 4:         'Nivel Millar
            If CInt(Mid(num, 1, 1)) = 1 Then
                Texto = "Mil " & LeerNumero(CStr(CInt(Mid(num, 2, 3))))
            Else
                Texto = Numero(CInt(Mid(num, 1, 1))) & " " & "mil" & " " & LeerNumero(CStr(CInt(Mid(num, 2, 3))))
            End If
    Case 5: '10.000 - 99.999
            Texto = LeerNumero(CStr(CInt(Mid(num, 1, 2)))) & " " & "mil" & " " & LeerNumero(CStr(CInt(Mid(num, 3, 3))))
    Case 6: '100.000 - 999.999
            Texto = LeerNumero(CStr(CInt(Mid(num, 1, 3)))) & " " & "mil" & " " & LeerNumero(CStr(CInt(Mid(num, 4, 3))))
    Case 7:        'Nivel Millón
            If CInt(Mid(num, 1, 1)) = 1 Then
                Texto = "Un Millón " & LeerNumero(CStr(CLng(Mid(num, 2, 6))))
            Else
                Texto = Numero(CInt(Mid(num, 1, 1))) & " " & "millones" & " " & LeerNumero(CStr(CLng(Mid(num, 2, 6))))
            End If
    Case 8 To 12: '10.000.000 - 999.999.999.999
            Texto = LeerNumero(CStr(CLng(Mid(num, 1, (level - 6))))) & " " & "millones" & " " & LeerNumero(CStr(CLng(Mid(num, (level - 5), 6))))
End Select

LeerNumero = Texto
Exit Function
error_leernumero:
    MsgBox "El Número está incorrectamente escrito", vbCritical, "Traductor Numérico"
    LeerNumero = "<número incorrecto>"
End Function
Public Function Mostrar_Reporte(Cnn As ADODB.Connection, Id As Integer, Fuente As String, Titulo As String, Opcion As Byte)
  Dim strRutaInformes As String
  strRutaInformes = GetSetting("Kit Logistics", "Configuracion", "RutaReportes")
  Dim rstReportes As New ADODB.Recordset
  rstReportes.CursorLocation = adUseClient
  rstReportes.Open "SELECT reportes.* FROM reportes WHERE IdReporte = " & Id, Cnn, adOpenDynamic, adLockOptimistic
    RutaInf = strRutaInformes & "\" & rstReportes.Fields("Tipo") & "\" & rstReportes.Fields("Ruta")
  rstReportes.Close
  
  If CpExisteFichero(RutaInf) = True Then
      TituloInf = Titulo
      rstFunPro.Open Fuente, Cnn, adOpenForwardOnly, adLockReadOnly
        OpcReporte = Opcion
        Dim frmRep As New FrmReportes
        frmRep.PasarParametros "Mario", CLng(1)
        frmRep.Show vbModal
        Set frmRep = Nothing
      rstFunPro.Close
  Else
    MsgBox "No existe el reporte " & RutaInf
  End If
End Function

Public Function SacarDatos(IdI As Integer, ArchivoKin As String) As String    'Modulok modulo para sacar el dato
  Dim i As Byte, FufuSt As String
  Open ArchivoKin For Input As #1
    For i = 0 To IdI
      Input #1, FufuSt, FufuSt, SacarDatos
    Next
  Close #1
End Function
Public Sub TiempoEspera(Milisegundos As Long)
  Sleep Milisegundos
End Sub

Public Function CpPermisoEspecial(IdPermiso As Integer, Usuario As Integer, Cnn As ADODB.Connection) As Boolean
  Dim rstUsuPermisos As New ADODB.Recordset
  rstUsuPermisos.CursorLocation = adUseClient
  rstUsuPermisos.Open "Select usupermisosesp.* from usupermisosesp  where IdUsuario=" & Usuario & " and IdPermiso=" & IdPermiso, Cnn, adOpenDynamic, adLockOptimistic
  If rstUsuPermisos.RecordCount > 0 Then
    CpPermisoEspecial = True
  Else
    CpPermisoEspecial = False
  End If
  rstUsuPermisos.Close
End Function

Public Function CpPermiso(IdFormulario As Integer, Usuario As Integer, Accion As Byte, Cnn As ADODB.Connection) As Boolean
  Dim rstPermisos As New ADODB.Recordset
  rstPermisos.CursorLocation = adUseClient
  rstPermisos.Open "Select permisos.* from permisos  where IdUsuario=" & Usuario & " and IdFormulario=" & IdFormulario, Cnn, adOpenDynamic, adLockOptimistic
  If rstPermisos.RecordCount > 0 Then
    Select Case Accion
      Case 1 'Ingreso
        If Val(rstPermisos.Fields("Ingreso")) = 1 Then
          CpPermiso = True
        Else
          MsgBox "El usuario no tiene permiso para ingresar a este modulo", vbCritical, "Permiso denegado"
          CpPermiso = False
        End If
      
      Case 2 'Nuevo
        If Val(rstPermisos.Fields("Nuevo")) = 1 Then
          CpPermiso = True
        Else
          MsgBox "El usuario no tiene permiso para crear registros en este modulo", vbCritical, "Permiso denegado"
          CpPermiso = False
        End If
      
      Case 3 'Editar
        If Val(rstPermisos.Fields("Editar")) = 1 Then
          CpPermiso = True
        Else
          MsgBox "El usuario no tiene permiso para editar en este modulo", vbCritical, "Permiso denegado"
          CpPermiso = False
        End If
      
      Case 4 'Anular o elimianr
        If Val(rstPermisos.Fields("Eliminar")) = 1 Then
          CpPermiso = True
        Else
          MsgBox "El usuario no tiene permiso para eliminar en este modulo", vbCritical, "Permiso denegado"
          CpPermiso = False
        End If
      
    End Select
  Else
    MsgBox "El usuario no tiene permisos asignados para este formulario", vbCritical, "Permiso denegado"
    CpPermiso = False
  End If
  rstPermisos.Close
End Function

Public Function EncryptString(UserKey As String, Text As String, Action As Single) As String
    Dim UserKeyX As String
    Dim Temp     As Integer
    Dim Times    As Integer
    Dim i        As Integer
    Dim j        As Integer
    Dim n        As Integer
    Dim rtn      As String
      
    '//Get UserKey characters
    n = Len(UserKey)
    ReDim UserKeyASCIIS(1 To n)
    For i = 1 To n
        UserKeyASCIIS(i) = Asc(Mid$(UserKey, i, 1))
    Next
          
    '//Get Text characters
    ReDim TextASCIIS(Len(Text)) As Integer
    For i = 1 To Len(Text)
        TextASCIIS(i) = Asc(Mid$(Text, i, 1))
    Next
      
    '//Encryption/Decryption
    If Action = ENCRYPT Then
       For i = 1 To Len(Text)
           j = IIf(j + 1 >= n, 1, j + 1)
           Temp = TextASCIIS(i) + UserKeyASCIIS(j)
           If Temp > 255 Then
              Temp = Temp - 255
           End If
           rtn = rtn + Chr$(Temp)
       Next
    ElseIf Action = DECRYPT Then
       For i = 1 To Len(Text)
           j = IIf(j + 1 >= n, 1, j + 1)
           Temp = TextASCIIS(i) - UserKeyASCIIS(j)
           If Temp < 0 Then
              Temp = Temp + 255
           End If
           rtn = rtn + Chr$(Temp)
       Next
    End If
      
    '//Return
    EncryptString = rtn
End Function

