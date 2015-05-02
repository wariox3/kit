VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAnalisisPendientesPordDespachar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Analisis por ruta de pendientes por despachar..."
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSelAll 
      Caption         =   "Seleccionar todo"
      Height          =   315
      Left            =   3480
      TabIndex        =   4
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CheckBox ChkAlInicio 
      Caption         =   "Ejecutar al iniciarse el sistema"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5880
      Width           =   2655
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   315
      Left            =   7800
      TabIndex        =   2
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton CmdAnalizar 
      Caption         =   "Analizar"
      Height          =   315
      Left            =   5520
      TabIndex        =   1
      Top             =   5880
      Width           =   2175
   End
   Begin MSComctlLib.ListView LstRutas 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   9975
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ruta"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Guias"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Unidades"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "K. Reales"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "K. Vol"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "U Desp. Viaje"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Dias"
         Object.Width           =   1058
      EndProperty
   End
End
Attribute VB_Name = "FrmAnalisisPendientesPordDespachar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAnalizar_Click()
  For II = 1 To LstRutas.ListItems.Count
    If LstRutas.ListItems(II).Checked = True Then
      AbrirRecorset rstUniversal, "SELECT COUNT(Guia) AS NumeroGuias, SUM(Unidades) AS TUND, SUM(KilosReales) AS TKR, SUM(KilosVolumen) AS TKV From Guias WHERE Anulada=0 AND Despachada=0 AND IdRuta = " & LstRutas.ListItems(II) & "", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      LstRutas.ListItems(II).SubItems(2) = Format(rstUniversal.Fields("NumeroGuias"), "#,##0;(#,##0)")
      If rstUniversal.Fields("NumeroGuias") > 0 Then
        LstRutas.ListItems(II).SubItems(3) = Format(rstUniversal.Fields("TUND"), "#,##0;(#,##0)")
        LstRutas.ListItems(II).SubItems(4) = Format(rstUniversal.Fields("TKR"), "#,##0;(#,##0)")
        LstRutas.ListItems(II).SubItems(5) = Format(rstUniversal.Fields("TKV"), "#,##0;(#,##0)")
      Else
        LstRutas.ListItems(II).SubItems(3) = 0
        LstRutas.ListItems(II).SubItems(4) = 0
        LstRutas.ListItems(II).SubItems(5) = 0
      End If
      CerrarRecorset rstUniversal
      AbrirRecorset rstUniversal, "SELECT MAX(FhExpedicion) AS UltDespacho From Despachos Where (IdRuta = " & LstRutas.ListItems(II) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
        LstRutas.ListItems(II).SubItems(6) = Format(rstUniversal.Fields("UltDespacho"), "dd/mm/yyyy")
        LstRutas.ListItems(II).SubItems(7) = Format(Date - rstUniversal.Fields("UltDespacho"), "#,##0;(#,##0)")
      CerrarRecorset rstUniversal
    End If
  Next
  
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub CmdSelAll_Click()
    For II = 1 To LstRutas.ListItems.Count
    If LstRutas.ListItems(II).Checked = False Then
      LstRutas.ListItems(II).Checked = True
    End If
  Next
End Sub

Private Sub Form_Load()
  ChkAlInicio.value = GetSetting("Kit Logistics", "Movimiento", "Inicio_Analiis_Rutas", 0)
  AbrirRecorset rstUniversal, "Select*From Rutas", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Do While rstUniversal.EOF = False
      Set Item = LstRutas.ListItems.Add(, , rstUniversal.Fields("IdRuta"))
        Item.SubItems(1) = rstUniversal.Fields("NmRuta")
        
      rstUniversal.MoveNext
    Loop
  CerrarRecorset rstUniversal
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "Kit logistics", "Movimiento", "Inicio_Analiis_Rutas", ChkAlInicio.value
End Sub
