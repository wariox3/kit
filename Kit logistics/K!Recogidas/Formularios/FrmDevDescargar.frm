VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmDevDescargar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Descargar recogida...."
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAyuda 
      Caption         =   "Ayuda"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPHora 
      Height          =   285
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "hh:mm"
      Format          =   16777218
      UpDown          =   -1  'True
      CurrentDate     =   38517
   End
   Begin VB.TextBox TxtKVol 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox TxtKReales 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox TxtUnidades 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   16777217
      CurrentDate     =   38517
   End
   Begin VB.Label LblTitulos 
      AutoSize        =   -1  'True
      Caption         =   "Fh Efectiva:"
      Height          =   195
      Index           =   4
      Left            =   2160
      TabIndex        =   12
      Top             =   480
      Width           =   855
   End
   Begin VB.Label LblTitulos 
      AutoSize        =   -1  'True
      Caption         =   "Hr Efectiva:"
      Height          =   195
      Index           =   3
      Left            =   2160
      TabIndex        =   11
      Top             =   120
      Width           =   840
   End
   Begin VB.Label LblTitulos 
      AutoSize        =   -1  'True
      Caption         =   "K Volumen:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   810
   End
   Begin VB.Label LblTitulos 
      AutoSize        =   -1  'True
      Caption         =   "K Reales:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   690
   End
   Begin VB.Label LblTitulos 
      AutoSize        =   -1  'True
      Caption         =   "Unidades:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "FrmDevDescargar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
  If Val(TxtUnidades.Text) > 0 Then
    If Val(TxtKReales.Text) > 0 Then
      If Val(TxtKVol.Text) > 0 Then
        FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(3) = Val(FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(3)) - Val(FrmVerRecogidas.LstAnuncios.ListItems(FrmVerRecogidas.LstAnuncios.SelectedItem.Index).SubItems(6))
        FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(4) = Val(FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(4)) - Val(FrmVerRecogidas.LstAnuncios.ListItems(FrmVerRecogidas.LstAnuncios.SelectedItem.Index).SubItems(7))
        FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(5) = Val(FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(5)) - Val(FrmVerRecogidas.LstAnuncios.ListItems(FrmVerRecogidas.LstAnuncios.SelectedItem.Index).SubItems(8))
        
        FrmVerRecogidas.LstAnuncios.ListItems(FrmVerRecogidas.LstAnuncios.SelectedItem.Index).SubItems(6) = Val(TxtUnidades)
        FrmVerRecogidas.LstAnuncios.ListItems(FrmVerRecogidas.LstAnuncios.SelectedItem.Index).SubItems(7) = Val(TxtKReales)
        FrmVerRecogidas.LstAnuncios.ListItems(FrmVerRecogidas.LstAnuncios.SelectedItem.Index).SubItems(8) = Val(TxtKVol)
        
        FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(3) = Val(FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(3)) + Val(FrmVerRecogidas.LstAnuncios.ListItems(FrmVerRecogidas.LstAnuncios.SelectedItem.Index).SubItems(6))
        FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(4) = Val(FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(4)) + Val(FrmVerRecogidas.LstAnuncios.ListItems(FrmVerRecogidas.LstAnuncios.SelectedItem.Index).SubItems(7))
        FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(5) = Val(FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(5)) + Val(FrmVerRecogidas.LstAnuncios.ListItems(FrmVerRecogidas.LstAnuncios.SelectedItem.Index).SubItems(8))
        AbrirRecorset rstUniversal, "Update Anuncios set Unidades=" & Val(TxtUnidades) & ", KilosReales=" & Val(TxtKReales) & ", KilosVol=" & Val(TxtKVol) & ", TiempoEfectiva='" & Format(DTPFecha, "dd/mm/yy") & " " & Format(DTPHora.Value, "hh:mm") & "', Efectiva=1, Estado='G' where IdAnuncio=" & FrmVerRecogidas.LstAnuncios.SelectedItem, CnnPrincipal, adOpenDynamic, adLockOptimistic
        FrmVerRecogidas.LstAnuncios.ListItems.Remove FrmVerRecogidas.LstAnuncios.SelectedItem.Index
        FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(3) = Val(FrmProgramarRecogidas.LstVehiculos.ListItems(FrmProgramarRecogidas.LstVehiculos.SelectedItem.Index).SubItems(3)) - 1
        Unload Me
      Else
        MsgBox "Los kilos volumen deben ser mayores a 0", vbCritical
      End If
    Else
      MsgBox "Los kilos reales deben ser mayores a 0", vbCritical
    End If
  Else
    MsgBox "Las unidades no pueden ser menores a 1", vbCritical
  End If
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub
Private Sub DTPFecha_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub
Private Sub DTPHora_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub
Private Sub TxtKReales_GotFocus()
  EnfocarT TxtKReales
End Sub
Private Sub TxtKReales_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtKReales, KeyAscii, 1
End Sub
Private Sub TxtKVol_GotFocus()
  EnfocarT TxtKVol
End Sub
Private Sub TxtKVol_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtKVol, KeyAscii, 1
End Sub
Private Sub TxtUnidades_GotFocus()
  EnfocarT TxtUnidades
End Sub
Private Sub TxtUnidades_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtUnidades, KeyAscii, 1
End Sub
