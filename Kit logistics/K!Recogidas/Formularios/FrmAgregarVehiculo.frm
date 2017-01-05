VERSION 5.00
Begin VB.Form FrmAgregarVehiculo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar vehiculo..."
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtIdConductor 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox TxtComentarios 
      Height          =   1005
      Left            =   960
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1200
      Width           =   6135
   End
   Begin VB.TextBox TxtRuta 
      Height          =   285
      Left            =   960
      MaxLength       =   5
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox TxtFlete 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox TxtVehiculo 
      Height          =   285
      Left            =   960
      MaxLength       =   6
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label NmConductor 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2400
      TabIndex        =   13
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Conductor:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   780
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Notas:"
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   11
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      Height          =   195
      Left            =   435
      TabIndex        =   10
      Top             =   840
      Width           =   390
   End
   Begin VB.Label LblConsulta 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   2400
      TabIndex        =   9
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Flete:"
      Height          =   195
      Index           =   1
      Left            =   5040
      TabIndex        =   8
      Top             =   120
      Width           =   390
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Vehiculo:"
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   7
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
    If FufuLo = 0 Then
      AbrirRecorset rstUniversal, "Insert into VehiculosRecogida (Fecha, Placa, Flete, Rec, Pend, Unidades, KilosReales, KilosVol, IdRuta, Notas, Coperaciones, UltOrden, IdConductor) VALUES(now(), '" & TxtVehiculo.Text & "', " & Val(TxtFlete.Text) & ", 0, 0, 0, 0, 0, " & Val(TxtRuta) & ", '" & TxtComentarios & "'," & Coperaciones & ",0, '" & TxtIdConductor.Text & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
    Else
      AbrirRecorset rstUniversal, "update vehiculosrecogida set Placa='" & TxtVehiculo.Text & "', Flete = " & Val(TxtFlete.Text) & ", IdRuta = " & Val(TxtRuta) & ", Notas = '" & TxtComentarios & "', IdConductor = '" & TxtIdConductor.Text & "' where IdAsignacion = " & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
    End If
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

Private Sub Form_Load()
  Dim strSql As String
  If FufuLo <> 0 Then
    strSql = "select vehiculosrecogida.* from vehiculosrecogida where IdAsignacion = " & FufuLo
    AbrirRecorset rstUniversal, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstUniversal.RecordCount > 0 Then
      TxtVehiculo.Text = rstUniversal.Fields("Placa")
      TxtIdConductor.Text = rstUniversal.Fields("IdConductor") & ""
      TxtFlete.Text = rstUniversal.Fields("Flete")
      TxtRuta.Text = rstUniversal.Fields("IdRuta")
    End If
  End If
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

Private Sub TxtIdConductor_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsulta 3, CnnPrincipal
    TxtIdConductor.Text = Principal.ToolConsultas1.DatSt
  End If
End Sub

Private Sub TxtIdConductor_LostFocus()
      If TxtIdConductor.Text <> "" Then
        AbrirRecorset rstUniversal, "Select IdConductor, Concat(Nombre, ' ', Apellido1,  ' ', Apellido2) as NmConductor, FhVenceLic, ConductorInactivo From Conductores where IdConductor='" & TxtIdConductor.Text & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          NmConductor.Caption = rstUniversal.Fields("NmConductor")
        Else
          MsgBox "El conductor no existe", vbCritical
          NmConductor.Caption = "": TxtIdConductor.Text = ""
        End If
        CerrarRecorset rstUniversal
      End If
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
