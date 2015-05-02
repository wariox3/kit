VERSION 5.00
Begin VB.Form FrmInfoCiudad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informacion de la la ciudad..."
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6015
      Begin VB.TextBox TxtIdDepartamento 
         Height          =   285
         Left            =   5280
         TabIndex        =   12
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox TxtIdCiudad 
         Height          =   285
         Left            =   1230
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtNmCiudad 
         Height          =   285
         Left            =   1230
         TabIndex        =   5
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox TxtDepartamento 
         Height          =   285
         Left            =   1230
         TabIndex        =   4
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox TxtCodMin 
         Height          =   285
         Left            =   1230
         TabIndex        =   3
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox TxtDistancia 
         Height          =   285
         Left            =   1230
         TabIndex        =   2
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Departamento:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   3
         Left            =   570
         TabIndex        =   10
         Top             =   600
         Width           =   600
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Index           =   4
         Left            =   630
         TabIndex        =   9
         Top             =   240
         Width           =   540
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Min:"
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   8
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Distancia:"
         Height          =   195
         Index           =   6
         Left            =   480
         TabIndex        =   7
         Top             =   1320
         Width           =   705
      End
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   120
      X2              =   6120
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      Index           =   0
      X1              =   6120
      X2              =   120
      Y1              =   2280
      Y2              =   2280
   End
End
Attribute VB_Name = "FrmInfoCiudad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  AbrirRecorset rstUniversal, "SELECT IdCiudad, NmCiudad, IdDepartamento, Distancia, IdZona, CodMinTrans  FROM Ciudades where IdCiudad=" & Val(FufuLo), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      TxtIdCiudad = rstUniversal!IdCiudad
      TxtNmCiudad = rstUniversal!NmCiudad & ""
      TxtIdDepartamento = rstUniversal!IdDepartamento & ""
      TxtDistancia = rstUniversal!Distancia & ""
      TxtCodMin = rstUniversal!CodMinTrans & ""
    End If
  CerrarRecorset rstUniversal
  If Val(TxtIdDepartamento) <> 0 Then TxtDepartamento = DevResBus("SELECT IdDepartamento, NmDepartamento From Departamentos where IdDepartamento=" & Val(TxtIdDepartamento), "NmDepartamento", CnnPrincipal)
End Sub
