VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReciboCaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibos Caja..."
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   7080
      TabIndex        =   22
      Top             =   2530
      Width           =   2295
      Begin VB.CheckBox ChkImpreso 
         Caption         =   "Impreso"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   975
      End
      Begin VB.CheckBox ChkAnulado 
         Caption         =   "Anulado"
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdVerDetalles 
      Caption         =   "Ver detalles"
      Height          =   255
      Left            =   7200
      TabIndex        =   21
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Frame FraAgregar 
      Enabled         =   0   'False
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   5640
      Width           =   6975
      Begin VB.CommandButton CmdRetirar 
         Caption         =   "Retirar marcados"
         Height          =   255
         Left            =   4920
         TabIndex        =   25
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   4920
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox TxtCuentaCobrar 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cobrar:"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Left            =   2760
         TabIndex        =   17
         Top             =   240
         Width           =   405
      End
   End
   Begin MSComctlLib.ListView LstReciboDet 
      Height          =   2415
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4260
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CxC"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Numero"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame FraEncabezado 
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   9255
      Begin VB.TextBox TxtCampos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   7560
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   7080
         TabIndex        =   20
         Top             =   240
         Width           =   405
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Index           =   6
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Recibo:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   9255
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtNmTercero 
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   6735
      End
      Begin VB.TextBox TxtCampos 
         Height          =   615
         Index           =   4
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   6735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tercero:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Comentarios:"
         Height          =   195
         Left            =   1440
         TabIndex        =   4
         Top             =   600
         Width           =   915
      End
   End
   Begin MSComctlLib.Toolbar ToolRecibos 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   1005
      ButtonWidth     =   847
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuev"
            Object.ToolTipText     =   "Crear nuevo registro [F9]"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Guar"
            Object.ToolTipText     =   "Guarda la informacio [F11]"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Object.ToolTipText     =   "Editar la informacion guardada [F10]"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Elim"
            Object.ToolTipText     =   "Elimina o anula el registro [F3]"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Can"
            Object.ToolTipText     =   "Cancela la creacion del registro [F4]"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Bus"
            Object.ToolTipText     =   "Buscar [Inicio]"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pri"
            Object.ToolTipText     =   "Ir al primer registro [F5]"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ant"
            Object.ToolTipText     =   "Ir al anterior registro [F6]"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sig"
            Object.ToolTipText     =   "Ir al siguiente registro [F7]"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ult"
            Object.ToolTipText     =   "Ir al ultimo registro [F8]"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cer"
            Object.ToolTipText     =   "Cerrar esta ventana [F12]"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Act"
            Object.ToolTipText     =   "Actualizar la informacion"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imp"
            Object.ToolTipText     =   "Imprimir registro [Fin]"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Car"
            Object.ToolTipText     =   "Carga informacion adicional [Pausa]"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmReciboCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstRecibos As New ADODB.Recordset
Dim strSqlRecibos As String
Dim Editando As Boolean

Private Sub CboTipo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub CmdAgregar_Click()
  If Val(TxtCuentaCobrar.Text) > 0 Then
    If Val(TxtValor.Text) > 0 Then
      Dim rstCuentaCobrar As New ADODB.Recordset
      rstCuentaCobrar.CursorLocation = adUseClient
      AbrirRecorset rstCuentaCobrar, "SELECT cuentas_cobrar.* FROM cuentas_cobrar WHERE IdCxC = " & Val(TxtCuentaCobrar.Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstCuentaCobrar.RecordCount > 0 Then
        If Val(TxtValor.Text) <= rstCuentaCobrar!Saldo Then
          AbrirRecorset rstUniversal, "INSERT INTO recibos_caja_det (IdRecibo, IdCxC, Valor) VALUES (" & Val(TxtCampos(0).Text) & ", " & Val(TxtCuentaCobrar.Text) & ", " & Val(TxtValor.Text) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
          AbrirRecorset rstUniversal, "UPDATE recibos_caja SET Total = Total + " & Val(TxtValor.Text) & " WHERE IdRecibo = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
          AbrirRecorset rstUniversal, "UPDATE cuentas_cobrar SET Saldo = Saldo - " & Val(TxtValor.Text) & " WHERE IdCxC = " & rstCuentaCobrar!IdCxC, CnnPrincipal, adOpenDynamic, adLockOptimistic
          LimpiarAgregar
          VerDetalle
          TxtCuentaCobrar.SetFocus
        Else
          MsgBox "El valor pagado es mayor al saldo de la cuenta por cobrar", vbCritical
          TxtValor.SetFocus
        End If
      End If
      CerrarRecorset rstCuentaCobrar
    Else
      MsgBox "El valor debe ser mayor a cero", vbCritical
      TxtValor.SetFocus
    End If
  Else
    MsgBox "Debe digitar una cuenta por cobrar", vbCritical
    TxtCuentaCobrar.SetFocus
  End If
End Sub

Private Sub CmdRetirar_Click()
  Dim rstReciboDet As New ADODB.Recordset
  rstReciboDet.CursorLocation = adUseClient
  II = 1
  While II <= LstReciboDet.ListItems.Count
    If LstReciboDet.ListItems(II).Checked = True Then
      AbrirRecorset rstReciboDet, "SELECT recibos_caja_det.* FROM recibos_caja_det WHERE IdReciboDet = " & LstReciboDet.SelectedItem, CnnPrincipal, adOpenDynamic, adLockOptimistic
        AbrirRecorset rstUniversal, "UPDATE recibos_caja SET Total = Total - " & rstReciboDet!Valor & " WHERE IdRecibo = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
        AbrirRecorset rstUniversal, "UPDATE cuentas_cobrar SET Saldo = Saldo + " & rstReciboDet!Valor & " WHERE IdCxC = " & rstReciboDet!IdCxC, CnnPrincipal, adOpenDynamic, adLockOptimistic
        AbrirRecorset rstUniversal, "DELETE FROM recibos_caja_det WHERE IdReciboDet=" & Val(LstReciboDet.SelectedItem), CnnPrincipal, adOpenDynamic, adLockOptimistic
      CerrarRecorset rstReciboDet
      LstReciboDet.ListItems.Remove (II)
    Else
     II = II + 1
    End If
    
  Wend
End Sub

Private Sub CmdVerDetalles_Click()
  VerDetalle
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  LaTecla (KeyCode), ToolRecibos
End Sub

Private Sub Form_Load()
  IconosTool ToolRecibos, Principal.IgListTool
  rstRecibos.CursorLocation = adUseServer
  
  strSqlRecibos = "SELECT recibos_caja.*, " & _
                "terceros.RazonSocial " & _
                "FROM recibos_caja " & _
                "LEFT JOIN terceros ON recibos_caja.IdTercero = terceros.IDTercero"
  AbrirRecorset rstRecibos, strSqlRecibos & " Order by IdRecibo Desc Limit 100", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Formatos rstRecibos
  Asignar rstRecibos
  Editando = False
End Sub

Private Sub Asignar(rstAsignar As ADODB.Recordset)
  For II = 0 To 4
    TxtCampos(II).Text = rstAsignar.Fields(II) & ""
  Next
  TxtNmTercero.Text = rstAsignar!RazonSocial & ""
  ChkImpreso.Value = DevCheck(rstAsignar!Impreso)
  ChkAnulado.Value = DevCheck(rstAsignar!Anulado)
  If Val(rstAsignar!Impreso) = 1 Or Val(rstAsignar!Anulado) = 1 Then
    FraAgregar.Enabled = False
    CmdRetirar.Enabled = False
  Else
    FraAgregar.Enabled = True
    If Editando = True Then
      CmdRetirar.Enabled = False
    End If
  End If
  If LstReciboDet.Tag = "Llena" Then
    LstReciboDet.ListItems.Clear
    LstReciboDet.Tag = "Vacia"
  End If
End Sub
Private Sub Formatos(rstForma As ADODB.Recordset)
  For II = 0 To 4
    Set rstForma.Fields(II).DataFormat = TxtCampos(II).DataFormat
  Next
End Sub
Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
        Desbloquear
        Limpiar
        TxtCampos(1) = Format(Date, "dd/mm/yy") & " " & Format(Time, "h:m:s")
        TxtCampos(2).SetFocus
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            Bloquear
            AbrirRecorset rstUniversal, "Update recibos_caja SET IdTercero='" & TxtCampos(2).Text & "', Total= " & Val(TxtCampos(3).Text) & ", Comentarios = '" & TxtCampos(4).Text & "'  WHERE IdRecibo = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
            Editando = False
            TxtCuentaCobrar.SetFocus
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
            TxtCampos(0).Text = SacarConsecutivo("RecibosCaja", CnnPrincipal)
            AbrirRecorset rstUniversal, "INSERT INTO recibos_caja (IdRecibo, Fecha, IdTercero, Total, Comentarios) " & _
                                        "VALUES (" & TxtCampos(0).Text & ", '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', '" & TxtCampos(2).Text & "', " & Val(TxtCampos(3).Text) & ", '" & TxtCampos(4).Text & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
            Bloquear
            AccionTool 17
            AccionTool 11
            TxtCuentaCobrar.SetFocus
        End If
      End If
    Case 5  'Editar
      If CpPermiso(3, CodUsuarioActivo, 3, CnnPrincipal) = True Then
          Editando = True
          Desbloquear
      End If
    Case 6 'Eliminar
      Dim rstActualizar As New ADODB.Recordset
      rstActualizar.CursorLocation = adUseClient
      AbrirRecorset rstUniversal, "SELECT recibos_caja.* FROM recibos_caja WHERE IdRecibo = " & Val(TxtCampos(0).Text) & " AND Impreso = 1 AND Anulado = 0", CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.RecordCount > 0 Then
        AbrirRecorset rstUniversal, "SELECT recibos_caja_det.* FROM recibos_caja_det WHERE IdRecibo = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
        Do While rstUniversal.EOF = False
          AbrirRecorset rstActualizar, "UPDATE cuentas_cobrar SET Saldo = Saldo + " & rstUniversal!Valor & " WHERE IdCxC = " & rstUniversal!IdCxC, CnnPrincipal, adOpenDynamic, adLockOptimistic
          rstUniversal.MoveNext
        Loop
        CerrarRecorset rstUniversal
        AbrirRecorset rstActualizar, "UPDATE recibos_caja_det SET Valor = 0 WHERE IdRecibo = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
        AbrirRecorset rstActualizar, "UPDATE recibos_caja SET Total = 0, Anulado = 1 WHERE IdRecibo = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
        MsgBox "Recibo anulado con exito", vbInformation
        AccionTool 17
        Asignar rstRecibos
      Else
        MsgBox "El recibo debe estar impreso y sin anular para poder ser anulado", vbCritical
      End If
      CerrarRecorset rstUniversal
    Case 7 'Cancelar
      If Editando = True Then
        AccionTool 4
      Else
        Asignar rstRecibos
        Bloquear
      End If
    Case 9  'Buscar
      If Principal.ToolConsultas1.AbrirDevDatos("Numero de recibo", "Digite el numero del recibo que desea buscar", 3, 0) = True Then
        AbrirRecorset rstUniversal, strSqlRecibos & " WHERE IdRecibo = " & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          Asignar rstUniversal
        Else
          MsgBox "No se encontraron recibos con este numero", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If
    Case 11 'Primero
      UPrimero rstRecibos
      Asignar rstRecibos
    Case 12 'Anterior
      UAnterior rstRecibos
      Asignar rstRecibos
    Case 13 'Siguiente
      USiguiente rstRecibos
      Asignar rstRecibos
    Case 14 'Ultimo
      UUltimo rstRecibos
      Asignar rstRecibos
    Case 16 'Cerrar
      Set rstRecibos = Nothing
      Unload Me
    Case 17 'Actualizar
      rstRecibos.Requery
      Formatos rstRecibos
    Case 18 'Imprimir
      AbrirRecorset rstUniversal, "SELECT recibos_caja.* FROM recibos_caja WHERE IdRecibo = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.RecordCount > 0 Then
        If Val(rstUniversal!Impreso) = 0 Then
          AbrirRecorset rstUniversal, "UPDATE recibos_caja SET Impreso = 1 WHERE IdRecibo = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
          Mostrar_Reporte CnnPrincipal, 31, "SELECT sql_ic_imprimir_recibo_caja.* FROM sql_ic_imprimir_recibo_caja WHERE IdRecibo = " & Val(TxtCampos(0).Text), "Imprimir recibo caja", 2
          AccionTool 17
          Asignar rstRecibos
        Else
          Mostrar_Reporte CnnPrincipal, 31, "SELECT sql_ic_imprimir_recibo_caja.* FROM sql_ic_imprimir_recibo_caja WHERE IdRecibo = " & Val(TxtCampos(0).Text), "Imprimir recibo caja", 2
        End If
      End If
      CerrarRecorset rstUniversal
    Case 19
    Case 20
  End Select
End Sub

Private Sub Desbloquear()
  BotTool 3, 17, ToolRecibos, True
  FraDatos.Enabled = True
  FraAgregar.Enabled = False
  CmdRetirar.Enabled = False
  CmdVerDetalles.Enabled = False
End Sub

Private Sub Bloquear()
  BotTool 3, 17, ToolRecibos, False
  FraDatos.Enabled = False
  FraAgregar.Enabled = True
  CmdVerDetalles.Enabled = True
  CmdRetirar.Enabled = True
End Sub

Private Sub Limpiar()
  For II = 0 To 4
    TxtCampos(II).Text = ""
  Next
  TxtNmTercero.Text = ""
End Sub
Private Sub LimpiarAgregar()
  TxtCuentaCobrar.Text = ""
  TxtValor.Text = ""
End Sub
Function Validacion() As Boolean
  If TxtCampos(2).Text <> "" Then
    Validacion = True
  Else
    Validacion = False: MsgTit "La factura debe tener un cliente": TxtCampos(2).SetFocus
  End If
End Function

Private Sub ToolRecibos_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub

Private Sub TxtCampos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Select Case Index
    Case 2
      If KeyCode = vbKeyF2 Then
        Principal.ToolConsultas1.AbrirDevConsulta 7, CnnPrincipal
        TxtCampos(2).Text = Principal.ToolConsultas1.DatSt
      End If
  End Select

End Sub

Private Sub TxtCampos_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
    Case 4
      If KeyAscii = 13 Then
        KeyAscii = 0
      End If
  End Select
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtCampos_LostFocus(Index As Integer)
  Select Case Index
    Case 2
      If Val(TxtCampos(2).Text) <> 0 Then
        AbrirRecorset rstUniversal, "SELECT IdTercero, RazonSocial FROM terceros WHERE IdTercero='" & TxtCampos(2).Text & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          TxtNmTercero = rstUniversal!RazonSocial & ""
        Else
          TxtNmTercero = "": TxtCampos(2).Text = ""
        End If
        CerrarRecorset rstUniversal
      End If
  End Select
End Sub

Private Sub TxtCuenta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtCuentaCobrar_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    FufuSt = Val(TxtCampos(2).Text)
    FrmAgregarCuentaCobrar.Show 1
    TxtCuentaCobrar.Text = FufuLo
  End If
End Sub

Private Sub TxtCuentaCobrar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtCuentaCobrar_Validate(Cancel As Boolean)
  AbrirRecorset rstUniversal, "SELECT cuentas_cobrar.* FROM cuentas_cobrar WHERE IdCxC = " & Val(TxtCuentaCobrar.Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.RecordCount > 0 Then
    TxtValor.Text = rstUniversal!Saldo
  End If
  CerrarRecorset rstUniversal
End Sub

Private Sub TxtValor_GotFocus()
EnfocarT TxtValor
End Sub

Private Sub TxtValor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub VerDetalle()
  Dim rstRecibosDetalle As New ADODB.Recordset
  rstRecibosDetalle.CursorLocation = adUseClient
  AbrirRecorset rstRecibosDetalle, "SELECT recibos_caja_det.*, NroDocumento, FechaDoc " & _
                          "FROM recibos_caja_det " & _
                          "LEFT JOIN cuentas_cobrar ON recibos_caja_det.IdCxC = cuentas_cobrar.IdCxC " & _
                          "WHERE IdRecibo = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
  LstReciboDet.ListItems.Clear
  If rstRecibosDetalle.RecordCount > 0 Then
    Do While rstRecibosDetalle.EOF = False
      Set Item = LstReciboDet.ListItems.Add(, , rstRecibosDetalle!IdReciboDet)
      Item.SubItems(1) = rstRecibosDetalle!IdCxC
      Item.SubItems(2) = rstRecibosDetalle!NroDocumento & ""
      Item.SubItems(3) = rstRecibosDetalle!FechaDoc & ""
      Item.SubItems(4) = rstRecibosDetalle!Valor
      rstRecibosDetalle.MoveNext
    Loop
  End If
  LstReciboDet.Tag = "Llena"
End Sub
