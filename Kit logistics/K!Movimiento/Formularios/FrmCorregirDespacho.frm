VERSION 5.00
Begin VB.Form FrmCorregirDespacho 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Corregir despacho"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   Icon            =   "FrmCorregirDespacho.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtCampos 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   9226
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   20
      Left            =   1260
      TabIndex        =   2
      Tag             =   "1"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox TxtCampos 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   9226
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   1275
      TabIndex        =   0
      Tag             =   "1"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox TxtNmCiudad 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   1995
      TabIndex        =   6
      Top             =   120
      Width           =   4695
   End
   Begin VB.TextBox TxtNmCiudad 
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   1995
      TabIndex        =   5
      Top             =   480
      Width           =   4695
   End
   Begin VB.TextBox TxtCampos 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   9226
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   1275
      TabIndex        =   1
      Tag             =   "1"
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label LblDespacho 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.Label LblUniversal 
      AutoSize        =   -1  'True
      Caption         =   "Anticipo:"
      Height          =   195
      Index           =   36
      Left            =   570
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Origen:"
      Height          =   195
      Left            =   675
      TabIndex        =   8
      Top             =   120
      Width           =   510
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Destino:"
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   480
      Width           =   585
   End
End
Attribute VB_Name = "FrmCorregirDespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstDespacho As New ADODB.Recordset

Private Sub CmdGuardar_Click()
  AbrirRecorset rstUniversal, "UPDATE despachos SET IdCiudadOrigen=" & TxtCampos(8).Text & ", IdCiudadDestino=" & TxtCampos(9).Text & ", VrAnticipo=" & TxtCampos(20).Text & " WHERE OrdDespacho = " & Val(LblDespacho.Caption), CnnPrincipal, adOpenDynamic, adLockOptimistic
  MsgBox "La correccion se realizo con exito", vbInformation
  Unload Me
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  rstDespacho.CursorLocation = adUseClient
  LblDespacho.Caption = FufuLo
  AbrirRecorset rstDespacho, "SELECT OrdDespacho, IdCiudadOrigen, IdCiudadDestino, ciudadOrigen.NmCiudad as NmCiudadOrigen, ciudadDestino.NmCiudad as NmCiudadDestino, VrAnticipo " & _
                             "FROM despachos " & _
                             "LEFT JOIN ciudades as ciudadOrigen ON despachos.IdCiudadOrigen = ciudadOrigen.IdCiudad " & _
                             "LEFT JOIN ciudades as ciudadDestino ON despachos.IdCiudadDestino = ciudadDestino.IdCiudad " & _
                             "WHERE OrdDespacho = " & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
  TxtCampos(8).Text = rstDespacho!IdCiudadOrigen
  TxtNmCiudad(8).Text = rstDespacho!NmCiudadOrigen
  TxtCampos(9).Text = rstDespacho!IdCiudadDestino
  TxtNmCiudad(9).Text = rstDespacho!NmCiudadDestino
  TxtCampos(20).Text = rstDespacho!VrAnticipo
End Sub

Private Sub TxtCampos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Select Case Index
      Case 8
        Principal.ToolConsultas1.AbrirDevConsulta 1, CnnPrincipal
        TxtCampos(8).Text = Principal.ToolConsultas1.DatLo
        
      Case 9
        Principal.ToolConsultas1.AbrirDevConsulta 1, CnnPrincipal
        TxtCampos(9).Text = Principal.ToolConsultas1.DatLo
      
    End Select
  End If
End Sub

Private Sub TxtCampos_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  Select Case Index
    Case 8, 9
      ValidarEntrada TxtCampos(Index), KeyAscii, 1
  End Select
End Sub

Private Sub TxtCampos_Validate(Index As Integer, Cancel As Boolean)
  Select Case Index
    Case 8, 9
      If Val(TxtCampos(Index).Text) <> 0 Then
        AbrirRecorset rstUniversal, "SELECT IdCiudad, NmCiudad, CodMinTrans  From Ciudades where IdCiudad=" & TxtCampos(Index), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          TxtNmCiudad(Index).Text = rstUniversal!NmCiudad & ""
          TxtCampos(Index).Tag = rstUniversal!CodMinTrans & ""
        Else
          TxtNmCiudad(Index).Text = "": TxtCampos(Index).Text = ""
        End If
        CerrarRecorset rstUniversal
      End If
  End Select
End Sub
