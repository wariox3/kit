VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmExportarManifiestos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar manifiestos"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "Agregar"
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox TxtOrdDespacho 
      Height          =   285
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.DirListBox DirArchivo 
      Height          =   4815
      Left            =   5880
      TabIndex        =   3
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton CmdExportar 
      Caption         =   "Exportar"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   7080
      TabIndex        =   1
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton CmdSeleccionarTodo 
      Caption         =   "Seleccionar todo"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   1455
   End
   Begin MSComctlLib.ListView LstDespachos 
      Height          =   4815
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   8493
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
         Text            =   "Orden"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Man"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ord Despacho:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "FrmExportarManifiestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rstVehiculos As New ADODB.Recordset
Dim rstTerceros As New ADODB.Recordset
Dim rstDespachosExp As New ADODB.Recordset



Private Sub CmdAgregar_Click()
  AbrirRecorset rstUniversal, "Select OrdDespacho, IdManifiesto, FhExpedicion from despachos where OrdDespacho=" & Val(TxtOrdDespacho.Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.RecordCount > 0 Then
    If rstUniversal.Fields("IdManifiesto") = 0 Then
      Set Item = LstDespachos.ListItems.Add(, , rstUniversal!OrdDespacho)
        Item.SubItems(1) = "99" & rstUniversal!OrdDespacho
        Item.SubItems(2) = Format(rstUniversal!FhExpedicion, "dd/mm/yy")
    Else
      MsgBox "El despacho tiene un manifiesto y no se puede agregar por esta opcion", vbCritical
    End If
  Else
    MsgBox "El despacho no existe", vbCritical
  End If
  CerrarRecorset rstUniversal
End Sub

Private Sub CmdExportar_Click()
  II = 1
  Dim ConceptoIndCom As String
  Open DirArchivo.Path & "\desexp" & Format(Date, "ddmmyy") & Format(Time, "hhmmss") & ".txt" For Append As #1
  IniProg 1, LstDespachos.ListItems.Count
  While II <= LstDespachos.ListItems.Count
    If LstDespachos.ListItems(II).Checked = True Then
      rstDespachosExp.Open "Select despachos.* from despachos, conductores where (despachos.idconductor=conductores.idconductor) and Exportado=0 and Estado<>'A' and OrdDespacho=" & LstDespachos.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstDespachosExp.RecordCount > 0 Then
        rstVehiculos.Open "Select IdPlaca, IdPropietario from vehiculos where IdPlaca='" & rstDespachosExp.Fields("IdVehiculo") & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
        rstTerceros.Open "select concat(Nombre, ' ', Apellido1, ' ', Apellido2 ) as NmTenedor from terceros where IdTercero='" & rstVehiculos.Fields("IdPropietario") & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
          Print #1, Format(rstDespachosExp.Fields("FhExpedicion"), "yyyymmdd") & "09" & Rellenar(rstDespachosExp.Fields("IdManifiesto"), 9, "0", 1) & Rellenar(rstVehiculos.Fields("IdPropietario"), 11, " ", 1) & "                                                  41450520  001NNLMR" & Rellenar(rstDespachosExp.Fields("IdManifiesto"), 9, "0", 1) & Rellenar(rstTerceros.Fields("NmTenedor"), 50, " ", 2) & "0001                                          " & Rellenar(rstDespachosExp.Fields("IdManifiesto"), 9, "0", 1) & "09N   001MANIFIESTO DE CARGA           99"
          Print #1, DevTpRetencion(Val(rstDespachosExp.Fields("VrFlete"))) & "99504   " & Rellenar(rstVehiculos.Fields("IdPropietario"), 11, " ", 2) & Rellenar(rstVehiculos.Fields("IdPropietario"), 11, " ", 1) & "00                            0             1.00" & Rellenar(Format(rstDespachosExp.Fields("VrFlete"), "##0.00;(##0.00)"), 17, " ", 1) & "  0.00  0.00             0.00             0.009999" & Rellenar(0#, 17, " ", 1) & "9999" & Rellenar(Format(rstDespachosExp.Fields("VrFlete"), "##0.00;(##0.00)"), 17, " ", 1) & "             0.00             0.00             0.00             1.00"
          Print #1, " N1NNNN              0.0099             0.00                            99             0.00                0.00      000000305             0.00             0.00    99N" & DevTpIndCom(Val(rstDespachosExp.Fields("VrFlete"))) & "S  0             0.00             0.00             0.00             0.00"
          Print #1, "             1.00                                                                    999999"
          'Credito
          Print #1, Format(rstDespachosExp.Fields("FhExpedicion"), "yyyymmdd") & "09" & Rellenar(rstDespachosExp.Fields("IdManifiesto"), 9, "0", 1) & Rellenar(rstVehiculos.Fields("IdPropietario"), 11, " ", 1) & "                                                  42505010  001NNLMR" & Rellenar(rstDespachosExp.Fields("IdManifiesto"), 9, "0", 1) & Rellenar(rstTerceros.Fields("NmTenedor"), 50, " ", 2) & "0001                                          " & Rellenar(rstDespachosExp.Fields("IdManifiesto"), 9, "0", 1) & "09N   001MANIFIESTO DE CARGA           99"
          Print #1, "9999504   " & Rellenar(rstVehiculos.Fields("IdPropietario"), 11, " ", 2) & Rellenar(rstVehiculos.Fields("IdPropietario"), 11, " ", 1) & "00                            0             1.00" & Rellenar(Format((Val(rstDespachosExp.Fields("VrFlete")) * 0.9) / 100, "##0.00;(##0.00)"), 17, " ", 1) & "  0.00  0.00             0.00             0.009999" & Rellenar(0#, 17, " ", 1) & "9999" & Rellenar(Format((Val(rstDespachosExp.Fields("VrFlete")) * 0.9) / 100, "##0.00;(##0.00)"), 17, " ", 1) & "             0.00             0.00             0.00             1.00"
          Print #1, " N2NNNN              0.0099             0.00                            99             0.00                0.00      000000305             0.00             0.00    99N99S  0             0.00             0.00             0.00             0.00"
          Print #1, "             1.00                                                                    999999"
        rstDespachosExp.Close
        rstVehiculos.Close
        rstTerceros.Close
        rstDespachosExp.Open "update despachos set exportado=1 where OrdDespacho=" & LstDespachos.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
      End If
      LstDespachos.ListItems.Remove (II)
      If rstDespachosExp.State <> adStateClosed Then
        rstDespachosExp.Close
      End If
    Else
     II = II + 1
    End If
    Prog (II)
  Wend
  FinProg
  Close #1
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub
Private Sub VerDespachos()
  LstDespachos.ListItems.Clear
  rstDespachosExp.Open "Select Despachos.* from Despachos where Exportado=0 and (Estado='I' or Estado='G' or Estado='V') and IdManifiesto<>0 order by IdManifiesto", CnnPrincipal, adOpenDynamic, adLockOptimistic
  IniProg 1, rstDespachosExp.RecordCount
  If rstDespachosExp.RecordCount > 0 Then
    Do While rstDespachosExp.EOF = False
      Prog (rstDespachosExp.AbsolutePosition)
      Set Item = LstDespachos.ListItems.Add(, , rstDespachosExp!OrdDespacho)
      Item.SubItems(1) = rstDespachosExp!IdManifiesto
      Item.SubItems(2) = Format(rstDespachosExp!FhExpedicion, "dd/mm/yy")
      Item.SubItems(3) = rstDespachosExp!VrFlete
      rstDespachosExp.MoveNext
    Loop
  End If
  FinProg
  rstDespachosExp.Close
End Sub

Private Sub CmdSeleccionarTodo_Click()
  II = 1
  For II = 1 To LstDespachos.ListItems.Count
    LstDespachos.ListItems(II).Checked = True
  Next
End Sub

Private Sub Form_Load()
  rstDespachosExp.CursorLocation = adUseClient
  rstVehiculos.CursorLocation = adUseClient
  rstTerceros.CursorLocation = adUseClient
  VerDespachos
  DirArchivo.Path = App.Path
End Sub

Private Function DevTpRetencion(Flete As Double) As String
  Dim RteFteMayor As Double
  AbrirRecorset rstUniversal, "Select RteFte, RteFteMayor, IndCom from ParametrizacionLiquidaciones", CnnPrincipal, adOpenDynamic, adLockOptimistic
    RteFteMayor = rstUniversal!RteFteMayor
  CerrarRecorset rstUniversal
  
  If Flete >= RteFteMayor Then
    DevTpRetencion = "01"
  Else
    DevTpRetencion = "99"
  End If
End Function

Private Function DevTpIndCom(Flete As Double) As String
  Dim RteFteMayor As Double
  AbrirRecorset rstUniversal, "Select RteFte, RteFteMayor, IndCom from ParametrizacionLiquidaciones", CnnPrincipal, adOpenDynamic, adLockOptimistic
    RteFteMayor = rstUniversal!RteFteMayor
  CerrarRecorset rstUniversal
  
  If Flete >= RteFteMayor Then
    DevTpIndCom = "06"
  Else
    DevTpIndCom = "99"
  End If
End Function

