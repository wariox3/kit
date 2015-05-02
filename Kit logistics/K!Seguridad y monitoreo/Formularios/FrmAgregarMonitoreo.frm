VERSION 5.00
Begin VB.Form FrmAgregarMonitoreo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar monitoreo..."
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox TxtDestino 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   5655
   End
   Begin VB.TextBox TxtIdVehiculo 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Destino:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vehiculo:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   660
   End
End
Attribute VB_Name = "FrmAgregarMonitoreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
  If TxtIdVehiculo.Text <> "" Then
    AbrirRecorset rstUniversal, "Insert into MonitoreoVehiculos (Orden, Tipo, Estado, Ok, FhHrSalida, Vehiculo, Destino, UltReporte, Frecuencia, EnNovedad, SinMonitoreo) Values (0, 5, 'P', 0, '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "','" & TxtIdVehiculo.Text & "', 'Recogida " & TxtDestino.Text & "', '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', 0, 0, 0)", CnnPrincipal, adOpenDynamic, adLockOptimistic
    MsgBox "Recogida elaborada con exito", vbInformation
    Unload Me
  Else
    MsgBox "Debe seleccionar un vehiculo", vbCritical
  End If
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub TxtDestino_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdVehiculo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsulta 5, CnnPrincipal
    TxtIdVehiculo.Text = Principal.ToolConsultas1.DatSt
  End If
End Sub

Private Sub TxtIdVehiculo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdVehiculo_Validate(Cancel As Boolean)
  AbrirRecorset rstUniversal, "Select IdPlaca from Vehiculos where IdPlaca='" & TxtIdVehiculo.Text & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.EOF = True Then
    MsgBox "El vehiculo no existe", vbCritical
    TxtIdVehiculo.Text = ""
  End If
  CerrarRecorset rstUniversal
End Sub
