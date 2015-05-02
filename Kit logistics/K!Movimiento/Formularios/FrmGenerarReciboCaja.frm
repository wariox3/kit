VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmGenerarReciboCaja 
   Caption         =   "Generar recibo caja"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   6375
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eliminar"
         Height          =   255
         Left            =   4800
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtVrManejo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   9
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox TxtVrFlete 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton CmdGenerar 
         Caption         =   "Generar"
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Manejo:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Flete:"
         Height          =   195
         Left            =   300
         TabIndex        =   10
         Top             =   240
         Width           =   390
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir recibo seleccionado"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4920
      Width           =   2415
   End
   Begin MSComctlLib.ListView LstRecibos 
      Height          =   2655
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "VrFlete"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "VrManejo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6375
      Begin VB.TextBox TxtGuia 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Guia:"
         Height          =   195
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   4920
      Width           =   2415
   End
End
Attribute VB_Name = "FrmGenerarReciboCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub CmdEliminar_Click()
  Dim rstActualizar As New ADODB.Recordset
  rstActualizar.CursorLocation = adUseClient
  If LstRecibos.ListItems.Count > 0 Then
    If MsgBox("Esta seguro de eliminar el recibo?", vbQuestion + vbYesNo) = vbYes Then
      AbrirRecorset rstUniversal, "SELECT ValorTotal FROM recibos_caja_soporte WHERE IdRecibo = " & Val(LstRecibos.SelectedItem), CnnPrincipal, adOpenDynamic, adLockOptimistic
      AbrirRecorset rstActualizar, "UPDATE guias SET Abonos=Abonos-" & Val(rstUniversal.Fields("ValorTotal")) & " WHERE Guia = " & Val(TxtGuia.Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
      AbrirRecorset rstUniversal, "DELETE FROM recibos_caja_soporte WHERE IdRecibo = " & Val(LstRecibos.SelectedItem), CnnPrincipal, adOpenDynamic, adLockOptimistic
      CerrarRecorset rstUniversal
      VerRecibos
    End If
  End If
End Sub

Private Sub CmdGenerar_Click()
  AbrirRecorset rstUniversal, "SELECT Guia, GuiFac, Estado, VrFlete, VrManejo, Abonos FROM guias WHERE Guia = " & Val(TxtGuia.Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
  If ((rstUniversal!VrFlete + rstUniversal!VrManejo) - rstUniversal!Abonos) < Val(TxtVrFlete.Text) + Val(TxtVrManejo.Text) Then
    MsgBox "El valor de los recibos no puede superar el valor del total de la guia", vbCritical
  Else
    AbrirRecorset rstUniversal, "INSERT INTO recibos_caja_soporte (FechaRecibo, VrFlete, VrManejo, ValorTotal, Guia) VALUES ('" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', " & Val(TxtVrFlete.Text) & ", " & Val(TxtVrManejo.Text) & ", " & Val(TxtVrFlete.Text) + Val(TxtVrManejo.Text) & ", " & Val(TxtGuia.Text) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
    AbrirRecorset rstUniversal, "UPDATE guias SET Abonos = Abonos + " & Val(TxtVrFlete.Text) + Val(TxtVrManejo.Text) & " WHERE Guia = " & Val(TxtGuia.Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
  End If
  VerRecibos
  
End Sub

Private Sub CmdImprimir_Click()
  If LstRecibos.ListItems.Count > 0 Then
    Mostrar_Reporte CnnPrincipal, 34, "SELECT sql_movimiento_formato_recibo_soporte.* FROM sql_movimiento_formato_recibo_soporte WHERE IdRecibo = " & Val(LstRecibos.SelectedItem), "Recibo caja soporte", 2
  End If
End Sub

Private Sub Form_Load()
  AbrirRecorset rstUniversal, "SELECT Guia, VrFlete, VrManejo, ExportadaCartera FROM guias WHERE Guia = " & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.RecordCount > 0 Then
    TxtGuia.Text = rstUniversal!Guia
    TxtVrFlete.Text = rstUniversal!VrFlete
    TxtVrManejo.Text = rstUniversal!VrManejo
    If Val(rstUniversal.Fields("ExportadaCartera")) = 1 Then
      CmdGenerar.Enabled = False
      CmdEliminar.Enabled = False
    End If
  End If
  CerrarRecorset rstUniversal
  VerRecibos
End Sub

Private Sub VerRecibos()
  LstRecibos.ListItems.Clear
  AbrirRecorset rstUniversal, "SELECT recibos_caja_soporte.* FROM recibos_caja_soporte WHERE Guia= " & Val(TxtGuia.Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    Set Item = LstRecibos.ListItems.Add(, , rstUniversal!IdRecibo)
    Item.SubItems(1) = rstUniversal!VrFlete
    Item.SubItems(2) = rstUniversal!VrManejo
    Item.SubItems(3) = rstUniversal!ValorTotal
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
End Sub
