VERSION 5.00
Begin VB.Form FrmVerImagen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ver Imagen"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   720
      Top             =   0
   End
   Begin VB.PictureBox PicOrigen 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PicImagen 
      Height          =   5895
      Left            =   120
      ScaleHeight     =   5835
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "FrmVerImagen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Timer1_Timer()
  If FufuLo <> 0 Then
    If CpExisteFichero(GetSetting("Kit Logistics", "Configuracion", "DirImagenes") & "conductores\" & FufuLo & ".jpg") = True Then
      PicOrigen = LoadPicture(GetSetting("Kit Logistics", "Configuracion", "DirImagenes") & "conductores\" & FufuLo & ".jpg")
      PicImagen.PaintPicture PicOrigen, 0, 0, PicImagen.ScaleWidth, PicImagen.ScaleHeight
    Else
      PicImagen.Picture = Nothing
    End If
  End If
  Timer1.Enabled = False
End Sub
