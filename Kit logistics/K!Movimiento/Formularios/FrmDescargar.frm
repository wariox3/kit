VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDescargar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Descargar..."
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   10230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkPermitirGuiasOtrosDespachos 
      Caption         =   "Permitir guias de otros despachos"
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   7440
      Width           =   3495
   End
   Begin VB.CommandButton CmdDescargarPorDocumento 
      Caption         =   "Descargar por documento"
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   6960
      Width           =   2415
   End
   Begin VB.CommandButton CmdPorRemision 
      Caption         =   "Descargar por guia"
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   6960
      Width           =   2415
   End
   Begin VB.CommandButton CmdSeleccionarTodas 
      Caption         =   "Seleccionar todas"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   6960
      Width           =   2415
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   7440
      Width           =   2415
   End
   Begin VB.CommandButton CmdDescargar 
      Caption         =   "Descargar seleccionadas"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   6960
      Width           =   2415
   End
   Begin MSComctlLib.ListView LstGuias 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   10610
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Destino"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ingreso"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Entregada"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Tipo"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Presione Ctr+P para ver las guias pendientes del despacho"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   7440
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   0
      X2              =   10080
      Y1              =   30
      Y2              =   30
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
      Left            =   2040
      TabIndex        =   7
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Manifiesto:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   765
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
      TabIndex        =   5
      Top             =   480
      Width           =   975
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
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Despacho:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   780
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      Index           =   1
      X1              =   10080
      X2              =   0
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Menu MnuAuxiliares 
      Caption         =   "&Reportes"
      Begin VB.Menu MnuPendientes 
         Caption         =   "Guias pendientes por descargar"
      End
      Begin VB.Menu MnuContraEntregas 
         Caption         =   "Cobros en el destino (Contra entregas)"
      End
   End
   Begin VB.Menu MnuAcciones 
      Caption         =   "A&cciones"
      Begin VB.Menu MnuVerPendientes 
         Caption         =   "Ver pendientes"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuVerMonitoreos 
         Caption         =   "Ver Monitoreos"
      End
   End
End
Attribute VB_Name = "FrmDescargar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Descargado As Boolean
Private Sub CmdDescargar_Click()
  Dim boolGuiasSinEntregar As Boolean
  Dim boolDescargarDestino As Boolean
  boolGuiasSinEntregar = False
  If CpPermisoEspecial(15, CodUsuarioActivo, CnnPrincipal) = True Then
    boolDescargarDestino = True
  End If
  II = 1
  While II <= LstGuias.ListItems.Count
    If LstGuias.ListItems(II).Checked = True Then
      If LstGuias.ListItems(II).SubItems(4) = "SI" Then
        If LstGuias.ListItems(II).SubItems(5) = 2 Then
          If boolDescargarDestino = True Then
            AbrirRecorset rstUniversal, "UPDATE Guias SET Descargada=1, Estado='G', FhDescargada= '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "' where Guia=" & Val(LstGuias.ListItems.Item(II)), CnnPrincipal, adOpenDynamic, adLockOptimistic
            InsertarLog 5, Val(LstGuias.ListItems.Item(II))
            LstGuias.ListItems.Remove (II)
          Else
            MsgBox "No tiene permiso para descargar guias destino", vbCritical
            II = II + 1
          End If
        Else
          AbrirRecorset rstUniversal, "UPDATE Guias SET Descargada=1, Estado='G', FhDescargada= '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "' where Guia=" & Val(LstGuias.ListItems.Item(II)), CnnPrincipal, adOpenDynamic, adLockOptimistic
          InsertarLog 5, Val(LstGuias.ListItems.Item(II))
          LstGuias.ListItems.Remove (II)
        End If
      Else
        boolGuiasSinEntregar = True
        II = II + 1
      End If
    Else
     II = II + 1
    End If
  Wend
  If boolGuiasSinEntregar = True Then
    MsgBox "Ha intentado descargar guias sin estado de entrega, verifique la entrega de estas guias", vbInformation
  End If
  If Descargado = False Then
    If LstGuias.ListItems.Count <= 0 Then
      Descargado = True
      DescargarDespacho
    End If
  End If
End Sub

Private Sub CmdDescargarPorDocumento_Click()
  Dim rstActGuia As New ADODB.Recordset
    Dim boolDescargarDestino As Boolean
    If CpPermisoEspecial(15, CodUsuarioActivo, CnnPrincipal) = True Then
      boolDescargarDestino = True
    End If
  rstActGuia.CursorLocation = adUseClient
  If Principal.ToolConsultas1.AbrirDevDatos("Numero de documento", "Digite el numero del documento de la guia que desea descargar", 2, 0) = True Then
    AbrirRecorset rstUniversal, "Select Guia, Entregada, TipoCobro from guias where IdDespacho = " & LblDespacho.Caption & " AND DocCliente='" & Principal.ToolConsultas1.DatSt & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstUniversal.RecordCount > 0 Then
      If Val(rstUniversal.Fields("Entregada")) = 1 Then
          If Val(rstUniversal.Fields("TipoCobro")) = 2 Then
            If boolDescargarDestino = True Then
              AbrirRecorset rstActGuia, "UPDATE Guias SET Descargada=1, Estado='G', FhDescargada= '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "' where Guia=" & rstUniversal.Fields("Guia"), CnnPrincipal, adOpenDynamic, adLockOptimistic
              InsertarLog 5, Val(rstUniversal.Fields("Guia"))
              CmdDescargarPorDocumento_Click
              VerPendientes
            Else
              MsgBox "No tiene permiso para descargar guias destino", vbCritical
            End If
          Else
            AbrirRecorset rstActGuia, "UPDATE Guias SET Descargada=1, Estado='G', FhDescargada= '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "' where Guia=" & rstUniversal.Fields("Guia"), CnnPrincipal, adOpenDynamic, adLockOptimistic
            InsertarLog 5, Val(rstUniversal.Fields("Guia"))
            CmdDescargarPorDocumento_Click
            VerPendientes
          End If
      End If
    Else
      MsgBox "No existen guias con este documento en este despacho", vbCritical
      CmdDescargarPorDocumento_Click
    End If
    CerrarRecorset rstUniversal
  End If
End Sub

Private Sub CmdPorRemision_Click()
  Dim rstActGuia As New ADODB.Recordset
  Dim boolDescargarDestino As Boolean
  If CpPermisoEspecial(15, CodUsuarioActivo, CnnPrincipal) = True Then
    boolDescargarDestino = True
  End If
  
  rstActGuia.CursorLocation = adUseClient
  If Principal.ToolConsultas1.AbrirDevDatos("Digite el numero de la guia", "Digite el numero de la guia que desea entregar", 3, 0) = True Then
    AbrirRecorset rstUniversal, "Select Guia, Entregada, IdDespacho, TipoCobro from guias where Guia=" & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstUniversal.RecordCount > 0 Then
      If ChkPermitirGuiasOtrosDespachos.value = 0 And rstUniversal.Fields("IdDespacho") <> Val(LblDespacho.Caption) Then
        MsgBox "La guia es de un despacho diferente y no esta habilitada la opcion de descargar guias de otros despachos"
      Else
        If Val(rstUniversal.Fields("Entregada")) = 1 Then
          If Val(rstUniversal.Fields("TipoCobro")) = 2 Then
            If boolDescargarDestino = True Then
              AbrirRecorset rstActGuia, "UPDATE Guias SET Descargada=1, Estado='G', FhDescargada= '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "' where Guia=" & rstUniversal.Fields("Guia"), CnnPrincipal, adOpenDynamic, adLockOptimistic
              InsertarLog 5, Val(rstUniversal.Fields("Guia"))
              CmdPorRemision_Click
              VerPendientes
            Else
              MsgBox "No tiene permiso para descargar guias destino", vbCritical
            End If
          Else
            AbrirRecorset rstActGuia, "UPDATE Guias SET Descargada=1, Estado='G', FhDescargada= '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "' where Guia=" & rstUniversal.Fields("Guia"), CnnPrincipal, adOpenDynamic, adLockOptimistic
            InsertarLog 5, Val(rstUniversal.Fields("Guia"))
            CmdPorRemision_Click
            VerPendientes
          End If
        End If
      End If
    Else
      MsgBox "No se encontro la guia", vbCritical
      CmdPorRemision_Click
    End If
    CerrarRecorset rstUniversal
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

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
  LblDespacho = FufuLo
  Descargado = False
End Sub

Sub VerPendientes()
  LstGuias.ListItems.Clear
  AbrirRecorset rstUniversalAux, "SELECT guias.Guia, guias.DocCliente, guias.FhEntradaBodega, guias.Estado, guias.Entregada, guias.IdDespacho, ciudades.NmCiudad, Descargada, TipoCobro FROM Guias INNER JOIN Ciudades ON Guias.IdCiuDestino = Ciudades.IdCiudad where Guias.IdDespacho=" & Val(LblDespacho) & " and Guias.Descargada=0", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversalAux.RecordCount > 0 Then
    'MsgBox "El despacho tiene " & rstUniversalAux.RecordCount & " guias pendientes por descargar", vbInformation, "Guias pendientes"
    IniProg 1, rstUniversalAux.RecordCount
    Do While rstUniversalAux.EOF = False
      Prog (rstUniversalAux.AbsolutePosition)
      Set Item = LstGuias.ListItems.Add(, , rstUniversalAux!Guia)
      Item.SubItems(1) = rstUniversalAux!DocCliente
      Item.SubItems(2) = rstUniversalAux!NmCiudad
      Item.SubItems(3) = rstUniversalAux!FhEntradaBodega
      Item.SubItems(5) = rstUniversalAux!TipoCobro
      If Val(rstUniversalAux!Entregada) = 1 Then
        Item.SubItems(4) = "SI"
      Else
        Item.SubItems(4) = "NO"
      End If
      rstUniversalAux.MoveNext
    Loop
    FinProg
  Else
    DescargarDespacho
  End If
  CerrarRecorset rstUniversalAux
End Sub
Sub DescargarDespacho()
  AbrirRecorset rstUniversal, "SELECT Estado FROM despachos WHERE OrdDespacho = " & Val(LblDespacho.Caption), CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.Fields("Estado") = "V" Then
    If MsgBox("Este despacho ya no tiene guias pendientes, se va a descargar automaticamente. ¿desea descargar el despado?", vbInformation + vbYesNo) = vbYes Then
      AbrirRecorset rstUniversal, "Update Despachos Set FhCumplidos='" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', Estado='G' where OrdDespacho=" & Val(LblDespacho), CnnPrincipal, adOpenDynamic, adLockBatchOptimistic
      MsgBox "El despacho fue descargado con exito, su nuevo estado es [DESCARGADO]", vbInformation, "Despacho descargado"
      Unload Me
    End If
  Else
  MsgBox "El despacho ya no tiene guias pendientes y esta descargado", vbInformation
    Unload Me
  End If
End Sub

Private Sub MnuContraEntregas_Click()
  Mostrar_Reporte CnnPrincipal, 8, "Select*from sql_im_contraentregas where IdDespacho=" & Val(LblDespacho.Caption), "Contraentregas y recaudos", 2
End Sub

Private Sub MnuPendientes_Click()
  Mostrar_Reporte CnnPrincipal, 12, "Select*from sql_im_pendescargar where IdDespacho=" & Val(LblDespacho.Caption), "Pendientes por descargar", 2
End Sub

Private Sub MnuVerMonitoreos_Click()
  FufuLo = Val(LblDespacho.Caption)
  FrmVerMonitoreos.Show 1
End Sub

Private Sub MnuVerPendientes_Click()
  VerPendientes
End Sub
