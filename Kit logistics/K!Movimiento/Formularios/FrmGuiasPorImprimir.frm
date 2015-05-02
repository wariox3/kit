VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmGuiasPorImprimir 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Guias por imprimir..."
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13755
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   13755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkPorImpresora 
      Caption         =   "Por impresora"
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton CmdMarcarTodas 
      Caption         =   "Marcar todas"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton CmdVerReporte 
      Caption         =   "Ver Reporte"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6480
      TabIndex        =   4
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton CmdActualizar 
      Caption         =   "Actualizar"
      Height          =   255
      Left            =   8880
      TabIndex        =   3
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton Cmdimprimir 
      Caption         =   "Imprimir marcadas"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   2055
   End
   Begin MSComctlLib.ListView LstGuias 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   7858
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Guia"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fh Entra"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cliente"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Remitente"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Documento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Destino"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Unidades"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "TCF"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "TCM"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar / Salir"
      Height          =   255
      Left            =   11280
      TabIndex        =   0
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label LblGuiasImpirmir 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nro de guias sin impirmir:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2145
   End
End
Attribute VB_Name = "FrmGuiasPorImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdAceptar_Click()
  Unload Me
End Sub

Private Sub VerGuias()
  LstGuias.ListItems.Clear
  AbrirRecorset rstUniversal, "SELECT Guia, FhEntradaBodega, Remitente, IdTpCtaFlete, IdTpCtaManejo, Cuenta, Cliente, DocCliente, IdCiuDestino, Unidades, Estado FROM Guias where Estado='D'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  LblGuiasImpirmir.Caption = rstUniversal.RecordCount
    Do While rstUniversal.EOF = False
      Set Item = LstGuias.ListItems.Add(, , rstUniversal.Fields("Guia"))
      Item.SubItems(1) = Format(rstUniversal.Fields("FhEntradaBodega"), "dd/mm/yy")
      Item.SubItems(2) = rstUniversal.Fields("Cuenta")
      Item.SubItems(3) = rstUniversal.Fields("Cliente")
      Item.SubItems(4) = rstUniversal.Fields("DocCliente")
      Item.SubItems(5) = rstUniversal.Fields("IdCiuDestino")
      Item.SubItems(6) = rstUniversal.Fields("Unidades")
      Item.SubItems(7) = rstUniversal.Fields("IdTpCtaFlete")
      Item.SubItems(8) = rstUniversal.Fields("IdTpCtaManejo")
      rstUniversal.MoveNext
    Loop
  CerrarRecorset rstUniversal
End Sub
Private Sub CmdActualizar_Click()
  VerGuias
End Sub

Private Sub CmdImprimir_Click()
  II = 1
  While II <= LstGuias.ListItems.Count
    If LstGuias.ListItems(II).Checked = True Then
      If ComprobarEstado(LstGuias.ListItems(II)) = "D" Then
        If ChkPorImpresora.value = 1 Then
          FufuLo = SelectForm("Rem Cuartas", Me.hwnd)
          ImprimirGuia LstGuias.ListItems(II).Text
        End If
        AbrirRecorset rstUniversal, "Update Guias Set Estado='I' where Guia=" & LstGuias.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
        LstGuias.ListItems.Remove (II)
      Else
        MsgBox "La guia " & LstGuias.ListItems(II) & " no esta en estado digitada", vbCritical
        LstGuias.ListItems.Remove (II)
      End If
    Else
     II = II + 1
    End If
  Wend
  
  On Error GoTo errPagina
    Printer.PaperSize = 1
errPagina:
  'If Err.Number = 380 Then
  '  MsgBox "Error " & Err.Number & " no esta configurada la impresora en el equipo", vbCritical
  'End If
End Sub

Private Sub CmdMarcarTodas_Click()
  For II = 1 To LstGuias.ListItems.Count
    LstGuias.ListItems(II).Checked = True
  Next
End Sub

Private Sub CmdVerReporte_Click()
    'Mostrar_Reporte CnnPrincipal, 11, "Select*from SQL_IM_PendientesImprimir", "", 2
End Sub

Private Sub Form_Load()
  VerGuias
End Sub
