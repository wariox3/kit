VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmAddCO 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar Centro de Operaciones..."
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtCO 
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo CboCiudades 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.TextBox TxtNombreCO 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ciudad:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "FrmAddCO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CboCiudades_GotFocus()
  LlenarCombo CboCiudades, "NmCiudad", "Ciudades"
End Sub

Private Sub CboCiudades_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub CboCiudades_LostFocus()
  AbrirRecorset rstUniversal, "Select IdCiudad, NmCiudad from ciudades where NmCiudad='" & CboCiudades.Text & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.EOF = False Then
    CboCiudades.Tag = rstUniversal!IdCiudad
  Else
    CboCiudades.Text = "": CboCiudades.Tag = ""
  End If
  CerrarRecorset rstUniversal
End Sub

Private Sub CmdAceptar_Click()
  If TxtNombreCO.Text <> "" Then
    If Val(CboCiudades.Tag) <> 0 Then
      Select Case II
        Case 1
          AbrirRecorset rstUniversal, "Insert into centrosoperaciones (NmPuntoOperaciones, IdCiudad, Tipo) values ('" & TxtNombreCO.Text & "', " & Val(CboCiudades.Tag) & ", 0)", CnnPrincipal, adOpenDynamic, adLockOptimistic
        Case 2
          AbrirRecorset rstUniversal, "Insert into centrosoperaciones (NmPuntoOperaciones, IdCiudad, Tipo) values ('" & TxtNombreCO.Text & "', " & Val(CboCiudades.Tag) & ", " & Val(TxtCO.Text) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
        Case 3
          AbrirRecorset rstUniversal, "Update centroscperaciones Set NmPuntoOperaciones='" & TxtNombreCO.Text & "', IdCiudad=" & Val(CboCiudades.Tag) & " where IdPO=" & Val(TxtCO.Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
      End Select
        MsgBox "Centro de Operaciones creado con exito", vbInformation
        Unload Me
    Else
      MsgBox "El Centro de Operaciones debe tener una ciudad", vbCritical
      CboCiudades.SetFocus
    End If
  Else
    MsgBox "El Centro de Operaciones debe tener un nombre", vbCritical
    TxtNombreCO.SetFocus
  End If
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  If II = 1 Then
    Me.Caption = "Crear centro de operaciones principal"
  Else
    Me.Caption = "Crear centro de operaciones secundario"
  End If
End Sub

Private Sub TxtNombreCO_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
