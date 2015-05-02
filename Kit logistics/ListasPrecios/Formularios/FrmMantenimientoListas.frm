VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMantenimientoListas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crear / Editar lista precios..."
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox TxtNombreLista 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPVence 
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Format          =   61341697
      CurrentDate     =   40742
   End
   Begin VB.Label LblIdLista 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Vence:"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "FrmMantenimientoListas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tipoEdicion As Integer
Private Sub CmdGuardar_Click()
  If TxtNombreLista.Text <> "" Then
    'Nuevo
    If tipoEdicion = 1 Then
      AbrirRecorset rstUniversal, "Insert into listasprecios (NmListaPrecios, FhVencimiento) Values('" & TxtNombreLista.Text & "', '" & Format(DTPVence.Value, "yy/mm/dd") & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
    Else
      AbrirRecorset rstUniversal, "UPDATE listasprecios SET FhVencimiento = '" & Format(DTPVence.Value, "yy/mm/dd") & "', NmListaPrecios ='" & TxtNombreLista.Text & "' WHERE IdListaPrecios = " & Val(LblIdLista.Caption), CnnPrincipal, adOpenDynamic, adLockOptimistic
    End If
    Unload Me
  Else
    MsgBox "Debe especificar un nombre para la lista", vbCritical
  End If
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  If FufuLo = 0 Then
    tipoEdicion = 1
    DTPVence = Date
  Else
    LblIdLista.Caption = FufuLo
    tipoEdicion = 2
    AbrirRecorset rstUniversal, "SELECT listasprecios.* FROM listasprecios WHERE IdListaPrecios = " & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
    TxtNombreLista.Text = rstUniversal!NmListaPrecios & ""
    DTPVence = rstUniversal!FhVencimiento
  End If
End Sub
