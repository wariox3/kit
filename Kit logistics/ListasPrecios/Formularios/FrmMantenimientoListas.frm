VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMantenimientoListas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crear / Editar lista precios..."
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtCodigoBufalo 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox TxtNombreLista 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPVence 
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Format          =   16842753
      CurrentDate     =   40742
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Codigo Bufalo:"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label LblIdLista 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6600
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Vence:"
      Height          =   195
      Left            =   885
      TabIndex        =   4
      Top             =   480
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      Height          =   195
      Left            =   795
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
      AbrirRecorset rstUniversal, "Insert into listasprecios (NmListaPrecios, FhVencimiento, codigo_empresa_bufalo) Values('" & TxtNombreLista.Text & "', '" & Format(DTPVence.Value, "yy/mm/dd") & "', '" & TxtCodigoBufalo.Text & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
    Else
      AbrirRecorset rstUniversal, "UPDATE listasprecios SET FhVencimiento = '" & Format(DTPVence.Value, "yy/mm/dd") & "', NmListaPrecios ='" & TxtNombreLista.Text & "', codigo_empresa_bufalo ='" & TxtCodigoBufalo.Text & "' WHERE IdListaPrecios = " & Val(LblIdLista.Caption), CnnPrincipal, adOpenDynamic, adLockOptimistic
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
    TxtCodigoBufalo.Text = rstUniversal.Fields("codigo_empresa_bufalo") & ""
    DTPVence = rstUniversal!FhVencimiento
  End If
End Sub
