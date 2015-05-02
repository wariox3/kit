VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCorreccionDigitosVerificacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Corregir digitos de verificacion"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSeleccionarTodos 
      Caption         =   "Seleccionar todos"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   2535
   End
   Begin VB.CommandButton CmdCorregirSeleccionados 
      Caption         =   "Corregir seleccionados"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   5760
      Width           =   2295
   End
   Begin MSComctlLib.ListView LstNits 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nit"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Actual"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Correcto"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   5760
      Width           =   1335
   End
End
Attribute VB_Name = "FrmCorreccionDigitosVerificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCorregirSeleccionados_Click()
    II = 1
  While II <= LstNits.ListItems.Count
    If LstNits.ListItems(II).Checked = True Then
      AbrirRecorset rstUniversal, "UPDATE terceros SET DigitoVerificacion = " & LstNits.ListItems(II).SubItems(3) & " WHERE IdTercero='" & LstNits.ListItems(II) & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
      LstNits.ListItems.Remove (II)
    Else
     II = II + 1
    End If
  Wend
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub CmdSeleccionarTodos_Click()
  II = 1
  While II <= LstNits.ListItems.Count
    LstNits.ListItems(II).Checked = True
    II = II + 1
  Wend
End Sub

Private Sub Form_Load()
  Dim rstNits As New ADODB.Recordset
  rstNits.CursorLocation = adUseClient
  AbrirRecorset rstNits, "SELECT terceros.* FROM terceros", CnnPrincipal, adOpenDynamic, adLockOptimistic
  While rstNits.EOF = False
    If Val(rstNits!DigitoVerificacion) <> DigitoVerificacion(rstNits!IdTercero) Then
      Set Item = LstNits.ListItems.Add(, , rstNits!IdTercero)
      Item.SubItems(1) = rstNits!RazonSocial & ""
      Item.SubItems(2) = rstNits!DigitoVerificacion & ""
      Item.SubItems(3) = DigitoVerificacion(rstNits!IdTercero)
    End If
    rstNits.MoveNext
  Wend
End Sub

Private Function DigitoVerificacion(ByVal sNit As String) As String
    On Error Resume Next
    Dim sTMP, sTmp1, sTmp2, aux As String
    Dim I As Integer
    Dim iResiduo  As Integer
    Dim iChequeo As Integer
    Dim iPrimos(15) As Integer '<- Defino el Arreglo de los Primos.
    For I = 1 To Len(sNit)
      If Mid(sNit, I, 1) <> "-" Then
        aux = aux & Mid(sNit, I, 1)
      End If
    Next I
    sNit = aux
    
    iPrimos(1) = 3: iPrimos(2) = 7: iPrimos(3) = 13: iPrimos(4) = 17: iPrimos(5) = 19
    iPrimos(6) = 23: iPrimos(7) = 29: iPrimos(8) = 37: iPrimos(9) = 41: iPrimos(10) = 43
    iPrimos(11) = 47: iPrimos(12) = 53: iPrimos(13) = 59: iPrimos(14) = 67: iPrimos(15) = 71
    iChequeo = 0: iResiduo = 0
    For I = 0 To Len(Trim(sNit)) - 1
        sTMP = Mid(sNit, Len(Trim(sNit)) - I, 1)
        iChequeo = iChequeo + (Val(sTMP) * iPrimos(I + 1))
        'MsgBox Val(sTmp), vbCritical, iPrimos(i + 1)
    Next I
    iResiduo = iChequeo Mod 11
    If iResiduo <= 1 Then
        If iResiduo = 0 Then DigitoVerificacion = 0
        If iResiduo = 1 Then DigitoVerificacion = 1
    Else
        DigitoVerificacion = 11 - iResiduo
    End If
    DigitoVerificacion = DigitoVerificacion
    'By GeNeTiKo
End Function
