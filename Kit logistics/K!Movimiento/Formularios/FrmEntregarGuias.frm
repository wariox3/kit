VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmEntregarGuias 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Entregar Guias"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkPermitirGuiasOtrosDespachos 
      Caption         =   "Permitir guias de otros despachos"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   8280
      Width           =   3495
   End
   Begin VB.CommandButton CmdEntregarPorDocumento 
      Caption         =   "Entregar por documento"
      Height          =   375
      Left            =   8040
      TabIndex        =   17
      Top             =   7800
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   7080
      Width           =   6735
      Begin VB.TextBox TxtHora 
         Height          =   285
         Left            =   3000
         TabIndex        =   12
         Text            =   "12"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TxtMin 
         Height          =   285
         Left            =   3960
         TabIndex        =   11
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox ChkDescargar 
         Caption         =   "Entregar y descargar"
         Height          =   255
         Left            =   4680
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DPFecha 
         Height          =   300
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   50003969
         CurrentDate     =   38971
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Hora:"
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Min:"
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton CmdPorRemision 
      Caption         =   "Entregar por guia"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CommandButton CmdSeleccionarTodas 
      Caption         =   "Seleccionar todas"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CommandButton CmdDescargar 
      Caption         =   "Entregar marcadas"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   8280
      Width           =   2535
   End
   Begin MSComctlLib.ListView LstGuias 
      Height          =   6615
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   11668
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Guia"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Documento"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Destino"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "FhIng"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nov"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cliente"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha de despacho:"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label LblFechaDespacho 
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
      Left            =   3840
      TabIndex        =   5
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Despacho:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FrmEntregarGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdDescargar_Click()
  Dim FechaServidor As Date
  II = 1
  If (Val(TxtHora.Text) > 0 And Val(TxtHora.Text) <= 23) And (Val(TxtMin.Text) >= 0 And Val(TxtMin.Text) <= 59) Then
    While II <= LstGuias.ListItems.Count
      If LstGuias.ListItems(II).Checked = True Then
        If CDate(LstGuias.ListItems(II).SubItems(3)) > CDate(DPFecha.value & " " & TxtHora.Text & ":" & TxtMin.Text) Then
          MsgBox "La fecha de despacho de la remision " & LstGuias.ListItems.Item(II) & " fue: [" & Format(LstGuias.ListItems(II).SubItems(3), "dd mmmm yyyy") & "] no puede descargar esta remision con una fecha inferior", vbCritical, "No se puede descargar"
          Exit Sub
        Else
          If CDate(LblFechaDespacho.Caption) > CDate(DPFecha.value & " " & TxtHora.Text & ":" & TxtMin.Text) Then
            MsgBox "La fecha del despacho no puede ser menor a la fecha de entrega", vbCritical
            Exit Sub
          Else
            AbrirRecorset rstUniversal, "Select now() as Fh", CnnPrincipal, adOpenDynamic, adLockOptimistic
            FechaServidor = rstUniversal.Fields("Fh")
            CerrarRecorset rstUniversal
            If CDate(FechaServidor) < CDate(DPFecha.value & " " & TxtHora.Text & ":" & TxtMin.Text) Then
              MsgBox "No puede entregar una guia con una fecha posterior a la actual"
              Exit Sub
            Else
              AbrirRecorset rstUniversal, "UPDATE Guias SET Entregada=1, Estado='G', FhEntregaMercancia= '" & Format(DPFecha.value, "yyyy/mm/dd") & " " & TxtHora.Text & ":" & TxtMin.Text & "', FhRegistroEntrega = '" & Format(Date, "yyyy/mm/dd") & "' where Guia=" & Val(LstGuias.ListItems.Item(II)), CnnPrincipal, adOpenDynamic, adLockOptimistic
              InsertarLog 4, Val(LstGuias.ListItems.Item(II))
              LstGuias.ListItems.Remove (II)
            End If
          End If
        End If
      Else
       II = II + 1
      End If
    Wend
  Else
    MsgBox "Hora no valida", vbCritical
  End If
End Sub

Private Sub CmdEntregarPorDocumento_Click()
Dim FechaServidor As Date
Dim rstDespacho As New ADODB.Recordset
rstDespacho.CursorLocation = adUseClient
Dim rstTemp As New ADODB.Recordset
rstTemp.CursorLocation = adUseClient
Dim rstAct As New ADODB.Recordset
rstAct.CursorLocation = adUseClient
  If Principal.ToolConsultas1.AbrirDevDatos("Numero de documento", "Digite el numero del documento de la guia que desea entregar", 2, 0) = True Then
    AbrirRecorset rstUniversal, "Select Guia, FhEntradaBodega, IdDespacho from guias where Despachada=1 and Anulada=0 and Entregada=0 AND IdDespacho=" & LblDespacho.Caption & " and DocCliente='" & Principal.ToolConsultas1.DatSt & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstUniversal.RecordCount > 0 Then
        If CDate(rstUniversal.Fields("FhEntradaBodega")) > CDate(DPFecha.value & " " & TxtHora.Text & ":" & TxtMin.Text) Then
          MsgBox "La fecha de despacho de la remision " & rstUniversal.Fields("Guia") & " fue: [" & Format(rstUniversal.Fields("FhEntradaBodega"), "dd mmmm yyyy") & "] no puede descargar esta remision con una fecha inferior", vbCritical, "No se puede descargar"
          Exit Sub
        Else
          AbrirRecorset rstDespacho, "select OrdDespacho, FhExpedicion from despachos where OrdDespacho=" & rstUniversal.Fields("IdDespacho"), CnnPrincipal, adOpenDynamic, adLockOptimistic
          If rstDespacho.RecordCount > 0 Then
            If CDate(rstDespacho.Fields("FhExpedicion")) > CDate(DPFecha.value & " " & TxtHora.Text & ":" & TxtMin.Text) Then
              MsgBox "La fecha del despacho no puede ser menor a la fecha de entrega", vbCritical
              Exit Sub
            Else
              AbrirRecorset rstTemp, "Select now() as Fh", CnnPrincipal, adOpenDynamic, adLockOptimistic
              FechaServidor = rstTemp.Fields("Fh")
              CerrarRecorset rstTemp
              If CDate(FechaServidor) < CDate(DPFecha.value & " " & TxtHora.Text & ":" & TxtMin.Text) Then
                MsgBox "No puede entregar una guia con una fecha posterior a la actual"
                Exit Sub
              Else
                AbrirRecorset rstAct, "UPDATE Guias SET Entregada=1, Estado='G', FhEntregaMercancia= '" & Format(DPFecha.value, "yyyy/mm/dd") & " " & TxtHora.Text & ":" & TxtMin.Text & "', FhRegistroEntrega = '" & Format(Date, "yyyy/mm/dd") & "' where Guia=" & Val(rstUniversal.Fields("Guia")), CnnPrincipal, adOpenDynamic, adLockOptimistic
                InsertarLog 4, Val(rstUniversal.Fields("Guia"))
                If ChkDescargar.value = 1 Then
                  AbrirRecorset rstAct, "UPDATE Guias SET Descargada=1, Estado='G', FhDescargada= '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "' where Guia=" & Val(rstUniversal.Fields("Guia")), CnnPrincipal, adOpenDynamic, adLockOptimistic
                  InsertarLog 5, Val(rstUniversal.Fields("Guia"))
                End If
              End If
            End If
          End If
          CerrarRecorset rstDespacho
        End If
    Else
      MsgBox "Verifique el numero del documento del cliente, que la guia no este anulada, que este despachada y que no este entregada, ademas que pertenezca a este despacho", vbCritical
    End If
    CerrarRecorset rstUniversal
    VerPendientes
    CmdEntregarPorDocumento_Click
  End If
End Sub

Private Sub CmdPorRemision_Click()
Dim FechaServidor As Date
Dim rstDespacho As New ADODB.Recordset
rstDespacho.CursorLocation = adUseClient
Dim rstTemp As New ADODB.Recordset
rstTemp.CursorLocation = adUseClient
Dim rstAct As New ADODB.Recordset
rstAct.CursorLocation = adUseClient
  If Principal.ToolConsultas1.AbrirDevDatos("Digite el numero de la guia", "Digite el numero de la guia que desea entregar", 3, 0) = True Then
    AbrirRecorset rstUniversal, "Select Guia, FhEntradaBodega, IdDespacho from guias where Despachada=1 and Anulada=0 and Entregada=0 AND Guia=" & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstUniversal.RecordCount > 0 Then
      If ChkPermitirGuiasOtrosDespachos.value = 0 And rstUniversal.Fields("IdDespacho") <> Val(LblDespacho.Caption) Then
        MsgBox "La guia no pertenece a este despacho y no esta habilitada la opcion de entregar guias de otros despachos"
      Else
        If CDate(rstUniversal.Fields("FhEntradaBodega")) > CDate(DPFecha.value & " " & TxtHora.Text & ":" & TxtMin.Text) Then
          MsgBox "La fecha de despacho de la remision " & rstUniversal.Fields("Guia") & " fue: [" & Format(rstUniversal.Fields("FhEntradaBodega"), "dd mmmm yyyy") & "] no puede descargar esta remision con una fecha inferior", vbCritical, "No se puede descargar"
          Exit Sub
        Else
          AbrirRecorset rstDespacho, "select OrdDespacho, FhExpedicion from despachos where OrdDespacho=" & rstUniversal.Fields("IdDespacho"), CnnPrincipal, adOpenDynamic, adLockOptimistic
          If rstDespacho.RecordCount > 0 Then
            If CDate(rstDespacho.Fields("FhExpedicion")) > CDate(DPFecha.value & " " & TxtHora.Text & ":" & TxtMin.Text) Then
              MsgBox "La fecha del despacho no puede ser menor a la fecha de entrega", vbCritical
              Exit Sub
            Else
              AbrirRecorset rstTemp, "Select now() as Fh", CnnPrincipal, adOpenDynamic, adLockOptimistic
              FechaServidor = rstTemp.Fields("Fh")
              CerrarRecorset rstTemp
              If CDate(FechaServidor) < CDate(DPFecha.value & " " & TxtHora.Text & ":" & TxtMin.Text) Then
                MsgBox "No puede entregar una guia con una fecha posterior a la actual"
                Exit Sub
              Else
                AbrirRecorset rstAct, "UPDATE Guias SET Entregada=1, Estado='G', FhEntregaMercancia= '" & Format(DPFecha.value, "yyyy/mm/dd") & " " & TxtHora.Text & ":" & TxtMin.Text & "', FhRegistroEntrega = '" & Format(Date, "yyyy/mm/dd") & "' where Guia=" & Val(rstUniversal.Fields("Guia")), CnnPrincipal, adOpenDynamic, adLockOptimistic
                InsertarLog 4, Val(rstUniversal.Fields("Guia"))
                If ChkDescargar.value = 1 Then
                  AbrirRecorset rstAct, "UPDATE Guias SET Descargada=1, Estado='G', FhDescargada= '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "' where Guia=" & Val(rstUniversal.Fields("Guia")), CnnPrincipal, adOpenDynamic, adLockOptimistic
                  InsertarLog 5, Val(rstUniversal.Fields("Guia"))
                End If
              End If
            End If
          End If
          CerrarRecorset rstDespacho
        End If
      End If
    Else
      MsgBox "Verifique el numero de la guia, que no este anulada, que este despachada y que no este entregada", vbCritical
    End If
    CerrarRecorset rstUniversal
    VerPendientes
    CmdPorRemision_Click
  End If
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub CmdSeleccionarTodas_Click()
  II = 1
  While II <= LstGuias.ListItems.Count
    LstGuias.ListItems(II).Checked = True
    II = II + 1
  Wend
End Sub

Private Sub DPFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then CmdDescargar.SetFocus
End Sub

Private Sub Form_Load()
  DPFecha.value = Date
  LblDespacho = FufuLo
  VerPendientes
End Sub

Sub VerPendientes()
  LstGuias.ListItems.Clear
  AbrirRecorset rstUniversalAux, "SELECT guias.Guia, guias.DocCliente, guias.FhEntradaBodega, guias.Estado, guias.IdDespacho, ciudades.NmCiudad, Descargada, EnNovedad, Cliente FROM Guias INNER JOIN Ciudades ON Guias.IdCiuDestino = Ciudades.IdCiudad where Guias.IdDespacho=" & Val(LblDespacho) & " and Guias.Entregada=0", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversalAux.RecordCount > 0 Then
    'MsgBox "El despacho tiene " & rstUniversalAux.RecordCount & " guias pendientes por entregar", vbInformation, "Guias pendientes"
    IniProg 1, rstUniversalAux.RecordCount
    Do While rstUniversalAux.EOF = False
      Prog (rstUniversalAux.AbsolutePosition)
      Set Item = LstGuias.ListItems.Add(, , rstUniversalAux!Guia)
      Item.SubItems(1) = rstUniversalAux!DocCliente
      Item.SubItems(2) = rstUniversalAux!NmCiudad
      Item.SubItems(3) = rstUniversalAux!FhEntradaBodega
      Item.SubItems(4) = rstUniversalAux!EnNovedad
      Item.SubItems(5) = rstUniversalAux!Cliente
      rstUniversalAux.MoveNext
    Loop
    FinProg
  End If
  CerrarRecorset rstUniversalAux
End Sub

Private Sub LstGuias_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then DPFecha.SetFocus
End Sub
