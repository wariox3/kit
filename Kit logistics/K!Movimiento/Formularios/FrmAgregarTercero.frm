VERSION 5.00
Begin VB.Form FrmAgregarTercero 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar tercero..."
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNmCliente 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   22
      Top             =   3120
      Width           =   6615
   End
   Begin VB.TextBox TxtNmCiudad 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   21
      Top             =   2760
      Width           =   6615
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   4560
      TabIndex        =   20
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "Agregar"
      Height          =   255
      Left            =   6600
      TabIndex        =   19
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox TxtCampo 
      Height          =   285
      Index           =   9
      Left            =   1080
      TabIndex        =   18
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox TxtCampo 
      Height          =   285
      Index           =   8
      Left            =   1080
      TabIndex        =   8
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox TxtCampo 
      Height          =   285
      Index           =   7
      Left            =   1080
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox TxtCampo 
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox TxtCampo 
      Height          =   285
      Index           =   6
      Left            =   1080
      TabIndex        =   6
      Top             =   2040
      Width           =   7335
   End
   Begin VB.TextBox TxtCampo 
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   4
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox TxtCampo 
      Height          =   285
      Index           =   4
      Left            =   1080
      TabIndex        =   3
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox TxtCampo 
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox TxtCampo 
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   5
      Top             =   1680
      Width           =   7335
   End
   Begin VB.TextBox TxtCampo 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Negociacion:"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   3120
      Width           =   945
   End
   Begin VB.Line Line1 
      X1              =   8520
      X2              =   120
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Ciudad:"
      Height          =   195
      Left            =   360
      TabIndex        =   17
      Top             =   2760
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Telefono:"
      Height          =   195
      Left            =   225
      TabIndex        =   16
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Direccion:"
      Height          =   195
      Left            =   180
      TabIndex        =   15
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Apellido2:"
      Height          =   195
      Left            =   210
      TabIndex        =   14
      Top             =   1320
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Apellido1:"
      Height          =   195
      Left            =   210
      TabIndex        =   13
      Top             =   960
      Width           =   690
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Extendido:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      Height          =   195
      Left            =   3240
      TabIndex        =   10
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   195
      Left            =   690
      TabIndex        =   9
      Top             =   240
      Width           =   210
   End
End
Attribute VB_Name = "FrmAgregarTercero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAgregar_Click()
  If TxtCampo(0).Text <> "" Then
    If TxtCampo(1).Text <> "" Then
      If (TxtCampo(1).Text = "C" Or TxtCampo(1).Text = "E") And (TxtCampo(3).Text = "" Or TxtCampo(4).Text = "") Then
        MsgBox "Por el tipo de identificacion el cliente debe tener por lo menos el nombre y un apellido", vbCritical: TxtCampo(3).SetFocus
      Else
        If TxtCampo(2).Text <> "" Then
          If TxtCampo(6).Text <> "" Then
            If TxtCampo(8).Text <> "" Then
              AbrirRecorset rstUniversal, "Select IdTercero From Terceros where IdTercero='" & TxtCampo(0) & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
              If rstUniversal.EOF = True Then
                CerrarRecorset rstUniversal
                AbrirRecorset rstUniversal, "INSERT INTO Terceros VALUES ('" & TxtCampo(0).Text & "', '" & TxtCampo(1).Text & "', '" & TxtCampo(2).Text & "', '" & TxtCampo(3).Text & "', '" & TxtCampo(4).Text & "', '" & TxtCampo(5).Text & "', '" & TxtCampo(6).Text & "', '" & TxtCampo(7).Text & "', " & TxtCampo(8) & ", '" & TxtCampo(9).Text & "', '',0)", CnnPrincipal, adOpenDynamic, adLockOptimistic
              Else
                CerrarRecorset rstUniversal
              End If
              FufuSt = TxtCampo(0)
              Unload Me
            Else
              MsgBox "El tercero debe tener una ciudad", vbCritical: TxtCampo(8).SetFocus
            End If
          Else
            MsgBox "El tercero debe tener una direccion", vbCritical: TxtCampo(6).SetFocus
          End If
        Else
          MsgBox "El tercero debe tener un nombre extendido", vbCritical: TxtCampo(2).SetFocus
        End If
      End If
    Else
      MsgBox "El tercero debe tener un tipo de identificacion", vbCritical: TxtCampo(1).SetFocus
    End If
  Else
    MsgBox "El tercero debe tener una identificacion", vbCritical: TxtCampo(0).SetFocus
  End If
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub
Private Sub Form_Load()
  TxtCampo(0).Text = FufuSt
End Sub

Private Sub TxtCampo_Change(Index As Integer)
  If Index >= 3 And Index <= 5 Then JuntarNombre
End Sub

Private Sub TxtCampo_GotFocus(Index As Integer)
  EnfocarT TxtCampo(Index)
  TxtCampo(Index).BackColor = &H80000001
  TxtCampo(Index).ForeColor = &HFFFFFF
End Sub

Private Sub TxtCampo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Select Case Index
      Case 8
        Principal.ToolConsultas1.AbrirDevConsulta 1, CnnPrincipal
        TxtCampo(8).Text = Principal.ToolConsultas1.DatLo
      Case 9
        Principal.ToolConsultas1.AbrirDevConsulta 2, CnnPrincipal
        TxtCampo(9).Text = Principal.ToolConsultas1.DatLo
    End Select
  End If
End Sub

Private Sub TxtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  Select Case Index
    Case 1
      ValidarEntrada TxtCampo(1), KeyAscii, 4
    Case 7, 8, 9
      ValidarEntrada TxtCampo(1), KeyAscii, 1
  End Select
End Sub

Private Sub TxtCampo_LostFocus(Index As Integer)
  TxtCampo(Index).BackColor = &H80000005
  TxtCampo(Index).ForeColor = &H80000012
End Sub

Private Sub TxtCampo_Validate(Index As Integer, Cancel As Boolean)
  Select Case Index
    Case 8
      If Val(TxtCampo(8).Text) <> 0 Then
        AbrirRecorset rstUniversal, "SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & TxtCampo(8), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          TxtNmCiudad.Text = rstUniversal!NmCiudad & ""
        Else
          TxtNmCiudad.Text = "": TxtCampo(8).Text = ""
        End If
        CerrarRecorset rstUniversal
      End If
      
    Case 9
      
      AbrirRecorset rstUniversal, "Select Id, NmNegociacion from Negociaciones where Id=" & Val(TxtCampo(9)), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        TxtNmCliente.Text = rstUniversal.Fields("NmNegociacion") & ""
      Else
        TxtNmCliente.Text = "": TxtCampo(9).Text = ""
      End If
      CerrarRecorset rstUniversal
  End Select
End Sub
Private Sub JuntarNombre()
  TxtCampo(2).Text = ""
  If TxtCampo(4).Text <> "" Then
    TxtCampo(2) = TxtCampo(2).Text & TxtCampo(4) & " "
  End If
  If TxtCampo(5).Text <> "" Then
    TxtCampo(2) = TxtCampo(2).Text & TxtCampo(5) & " "
  End If
  If TxtCampo(3).Text <> "" Then
    TxtCampo(2) = TxtCampo(2).Text & TxtCampo(3)
  End If
End Sub
