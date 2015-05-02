VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmContraseña 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingreso al sistema..."
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "FrmContraseña.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContraseña.frx":08CA
            Key             =   "Usu"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ayuda"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox TxtContraseña 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin MSComctlLib.ImageCombo CboUsuarios 
      Height          =   330
      HelpContextID   =   6
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Text            =   "                    "
      ImageList       =   "ImageList2"
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   4560
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.Label LblUsuario 
      AutoSize        =   -1  'True
      Caption         =   "Usuario:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   585
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   120
      X2              =   4560
      Y1              =   960
      Y2              =   960
   End
End
Attribute VB_Name = "FrmContraseña"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstUsuarios As New ADODB.Recordset
Const ENCRYPT = 1
Const DECRYPT = 2
Private Sub CboUsuarios_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub CmdAceptar_Click()
  Dim Criterio As String
  Select Case ModIngreso
    Case 1 'Movimiento
      Criterio = " And ModMovimiento=1"
    Case 2 'Monitoreo
      Criterio = " And ModMonitoreo=1"
    Case 3 'Facturacion
      Criterio = " And ModFacturacion=1"
    Case 4 'Recogidas
      Criterio = " And ModRecogidas=1"
    Case 5 'Vechiculos
      Criterio = " And ModVehiculos=1"
    Case 6 'DatosBasicos
      Criterio = " And ModDatosBasicos=1"
  End Select
  
  rstUsuarios.Open "select*from usuarios where NmUsuario='" & CboUsuarios.Text & "' and Contraseña='" & EncryptString("mario", TxtContraseña.Text, 1) & "'" & Criterio, CnnSeguridad, adOpenStatic, adLockReadOnly
  If rstUsuarios.EOF = False Then
      SgLo = rstUsuarios!IdUsuario
      Unload Me
  Else
    MsgBox "Nombre de usuario o contraseña incorrectos, verifique si tiene permisos para el ingreso a este modulo ", vbCritical, "Usuario Incorrecto"
    CboUsuarios.SetFocus
  End If
  rstUsuarios.Close
End Sub
Private Sub CmdCancelar_Click()
  Unload Me
  SgLo = 0
End Sub
Private Sub Form_Load()
  SgLo = 0
  CboUsuarios.ComboItems.Clear
  rstUsuarios.Open "select * from usuarios where Inactivo=0 order by NmUsuario", CnnSeguridad, adOpenDynamic, adLockOptimistic
  While Not rstUsuarios.EOF = True
    CboUsuarios.ComboItems.Add , , rstUsuarios.Fields("NmUsuario"), "Usu", "Usu"
    rstUsuarios.MoveNext
  Wend
  rstUsuarios.Close
End Sub
Private Sub TxtContraseña_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then CmdAceptar_Click
End Sub

Private Function EncryptString(UserKey As String, Text As String, Action As Single) As String
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

