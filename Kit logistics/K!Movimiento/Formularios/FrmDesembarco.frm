VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDesembarco 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Des-embarcar viaje..."
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdMarcarTodas 
      Caption         =   "Marcar todas"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton CmdDesembarcarMarcadas 
      Caption         =   "Desembarcar Marcadas"
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   5640
      Width           =   2295
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   255
      Left            =   5520
      TabIndex        =   0
      Top             =   5640
      Width           =   1815
   End
   Begin MSComctlLib.ListView LstGuias 
      Height          =   4695
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Guia"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Destino"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "FhIng"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Despacho:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   780
   End
   Begin VB.Label LblDespacho 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label LblManifiesto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Manifiesto:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   765
   End
   Begin VB.Label LblEstado 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "FrmDesembarco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdDesembarcarMarcadas_Click()
  II = 1
  While II <= LstGuias.ListItems.Count
    If LstGuias.ListItems(II).Checked = True Then
      AbrirRecorset rstUniversal, "UPDATE Guias SET Estado='E', IdDespacho = NULL, Despachada = 0, CR=" & Coperaciones & "  where Guia=" & Val(LstGuias.ListItems.Item(II)), CnnPrincipal, adOpenDynamic, adLockOptimistic
      InsertarLog 9, Val(LstGuias.ListItems.Item(II))
      LstGuias.ListItems.Remove (II)
    Else
     II = II + 1
    End If
  Wend
  LstGuias.SetFocus
End Sub

Private Sub CmdMarcarTodas_Click()
  II = 1
  While II <= LstGuias.ListItems.Count
    LstGuias.ListItems(II).Checked = True
    II = II + 1
  Wend
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  LblDespacho.Caption = FufuLo
  VerGuias
End Sub
Sub VerGuias()
  LstGuias.ListItems.Clear
  AbrirRecorset rstUniversal, "SELECT Guias.Guia, Guias.FhEntradaBodega, Guias.Estado, Guias.IdDespacho, Ciudades.NmCiudad FROM Guias INNER JOIN Ciudades ON Guias.IdCiuDestino = Ciudades.IdCiudad where Guias.IdDespacho=" & Val(LblDespacho) & " and Guias.Estado<>'G' and Guias.Estado<>'E'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    MsgBox "El despacho tiene " & rstUniversal.RecordCount & " guias para desembarcar", vbInformation, "Guias"
    IniProg 1, rstUniversal.RecordCount
    Do While rstUniversal.EOF = False
      Prog (rstUniversal.AbsolutePosition)
      Set Item = LstGuias.ListItems.Add(, , rstUniversal!Guia)
      Item.SubItems(1) = rstUniversal!NmCiudad
      Item.SubItems(2) = rstUniversal!FhEntradaBodega
      rstUniversal.MoveNext
    Loop
    FinProg
  CerrarRecorset rstUniversal
End Sub
