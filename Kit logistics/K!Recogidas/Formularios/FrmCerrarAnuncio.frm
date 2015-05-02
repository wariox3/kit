VERSION 5.00
Begin VB.Form FrmCerrarAnuncio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cerrar Anuncio"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtMotivoCerrarRecogida 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   8175
   End
   Begin VB.CommandButton CmdCerrarRecogida 
      Caption         =   "Cerrar o cancelar recogida"
      Height          =   255
      Left            =   5640
      TabIndex        =   0
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Especifique los motivos por los cuales desea cerrar o cancelar esta recogida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "FrmCerrarAnuncio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCerrarRecogida_Click()
  If MsgBox("¿Esta seguro de cerrar el anuncio de recogida?", vbYesNo + vbQuestion) = vbYes Then
    AbrirRecorset rstUniversal, "UPDATE anuncios SET Cerrada = 1, Estado = 'C', MotivoCancelacion = '" & TxtMotivoCerrarRecogida.Text & "'  where IdAnuncio=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockBatchOptimistic
    MsgBox "El anuncio de recogida ha sido cerrado con exito", vbInformation
    Unload Me
  End If
End Sub
