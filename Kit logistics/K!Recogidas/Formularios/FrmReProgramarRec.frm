VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReProgramarRec 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Re-Programar recogida..."
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPHora 
      Height          =   300
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16908290
      CurrentDate     =   38516
   End
   Begin MSComCtl2.DTPicker DTPFechaRe 
      Height          =   300
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   16908291
      CurrentDate     =   38516
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Hora:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   390
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "FrmReProgramarRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
  If DTPFechaRe.Value >= Date Then
    AbrirRecorset rstUniversal, "Update Anuncios set FhRecogida='" & Format(DTPFechaRe.Value, "yy/mm/dd") & " " & Format(DTPHora.Value, "h:m:s") & "' where IdAnuncio=" & FrmProgramarRecogidas.LstAnuncios.SelectedItem, CnnPrincipal, adOpenDynamic, adLockOptimistic
    MsgBox "La recogida fue reprogramamda para el " & Format(DTPFechaRe.Value, "dd mmmm yyyy") & " a las " & Format(DTPHora.Value, "hh:mm"), vbInformation
    Unload Me
  Else
    MsgBox "Solo puede re-programar esta recogida para un dia posterior o igual al actual", vbCritical
  End If
End Sub
Private Sub CmdCancelar_Click()
  Unload Me
End Sub
Private Sub DTPFechaRe_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub
Private Sub DTPHora_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub
