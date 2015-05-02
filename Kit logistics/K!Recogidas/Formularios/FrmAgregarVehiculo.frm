VERSION 5.00
Begin VB.Form FrmAgregarVehiculo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar vehiculo..."
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtComentarios 
      Height          =   1005
      Left            =   840
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   840
      Width           =   5655
   End
   Begin VB.TextBox TxtRuta 
      Height          =   285
      Left            =   840
      MaxLength       =   5
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox TxtFlete 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox TxtVehiculo 
      Height          =   285
      Left            =   840
      MaxLength       =   6
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Notas:"
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   10
      Top             =   840
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      Height          =   195
      Left            =   435
      TabIndex        =   9
      Top             =   480
      Width           =   390
   End
   Begin VB.Label LblConsulta 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   8
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Flete:"
      Height          =   195
      Index           =   1
      Left            =   4440
      TabIndex        =   7
      Top             =   120
      Width           =   390
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Vehiculo:"
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   6
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "FrmAgregarVehiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
If TxtVehiculo.Text <> "" Then
  If Val(TxtRuta) <> 0 Then
    II = 1
    AbrirRecorset rstUniversal, "Insert into VehiculosRecogida (Fecha, Placa, Flete, Rec, Pend, Unidades, KilosReales, KilosVol, IdRuta, Notas, Coperaciones, UltOrden) VALUES(now(), '" & TxtVehiculo.Text & "', " & Val(TxtFlete.Text) & ", 0, 0, 0, 0, 0, " & Val(TxtRuta) & ", '" & TxtComentarios & "'," & Coperaciones & ",0)", CnnPrincipal, adOpenDynamic, adLockOptimistic
    MsgBox "Vehiculo agregado con exito", vbInformation
    Unload Me
  Else
    MsgBox "El vehiculo debe llevar una ruta", vbCritical
    TxtRuta.SetFocus
  End If
Else
  II = 0
  MsgBox "Debe digitar un vehiculo, en el caso de que no se sepa la placa, ubiquese en el campo de vehiculo y presione F2", vbCritical
  TxtVehiculo.SetFocus
End If
End Sub

Private Sub CmdAgregarAuxiliar_Click()
  Principal.ToolConsultas1.AbrirDevConsulta 10, CnnPrincipal
  If Principal.ToolConsultas1.DatSt <> "" <> "" Then
    
  End If
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub TxtComentarios_GotFocus()
  EnfocarT TxtComentarios
End Sub

Private Sub TxtComentarios_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys vbTab
  End If
End Sub
Private Sub TxtFlete_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtFlete, KeyAscii, 1
End Sub

Private Sub TxtRuta_GotFocus()
  EnfocarT TxtRuta
End Sub

Private Sub TxtRuta_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsultaCO 1, Coperaciones, CnnPrincipal
    TxtRuta.Text = Principal.ToolConsultas1.DatLo
  End If
End Sub

Private Sub TxtRuta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtRuta_Validate(Cancel As Boolean)
  If Val(TxtRuta) <> 0 Then
    AbrirRecorset rstUniversal, "SELECT IdRutaRec, NmRuta FROM RutasUrbanas where IdRutaRec=" & Val(TxtRuta), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      LblConsulta(2) = rstUniversal!NmRuta & ""
    Else
      LblConsulta(2) = "": TxtRuta = ""
    End If
    CerrarRecorset rstUniversal
  Else
    LblConsulta(2) = ""
  End If
End Sub

Private Sub TxtVehiculo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsulta 5, CnnPrincipal
    TxtVehiculo.Text = Principal.ToolConsultas1.DatSt
  End If
End Sub
Private Sub TxtVehiculo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub TxtVehiculo_Validate(Cancel As Boolean)
  AbrirRecorset rstUniversal, "Select IdPlaca from vehiculos where IdPlaca='" & TxtVehiculo.Text & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.RecordCount < 1 Then TxtVehiculo.Text = ""
  CerrarRecorset rstUniversal
End Sub
