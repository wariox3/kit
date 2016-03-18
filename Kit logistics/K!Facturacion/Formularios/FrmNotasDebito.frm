VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmNotasDebito 
   Caption         =   "Notas debito..."
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   7080
      TabIndex        =   19
      Top             =   2530
      Width           =   2295
      Begin VB.CheckBox ChkImpreso 
         Caption         =   "Impreso"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   975
      End
      Begin VB.CheckBox ChkAnulado 
         Caption         =   "Anulado"
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdVerDetalles 
      Caption         =   "Ver detalles"
      Height          =   255
      Left            =   7200
      TabIndex        =   17
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Frame FraAgregar 
      Enabled         =   0   'False
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   5640
      Width           =   6855
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   4920
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TxtCuentaCobrar 
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton CmdRetirar 
         Caption         =   "Retirar marcados"
         Height          =   255
         Left            =   4920
         TabIndex        =   14
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cobrar:"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Left            =   2760
         TabIndex        =   15
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.Frame FraEncabezado 
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   9255
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   5
         Left            =   2520
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtCampos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   7560
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   1
         Left            =   4560
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Numero:"
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   27
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   7080
         TabIndex        =   12
         Top             =   240
         Width           =   405
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Index           =   6
         Left            =   3960
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   210
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
      Begin VB.Label TxtNmTercero 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   6735
      End
   End
   Begin MSComctlLib.ListView LstNotaDebitoDet 
      Height          =   2415
      Left            =   120
      TabIndex        =   24
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
   Begin MSComctlLib.Toolbar ToolNotasDebito 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   9630
      _ExtentX        =   16986
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
Attribute VB_Name = "FrmNotasDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rstNotasDebito As New ADODB.Recordset
Dim strSqlNotasDebito As String
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
          AbrirRecorset rstUniversal, "INSERT INTO notas_debito_det (IdNotaDebito, IdCxC, Valor) VALUES (" & Val(TxtCampos(0).Text) & ", " & Val(TxtCuentaCobrar.Text) & ", " & Val(TxtValor.Text) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
          AbrirRecorset rstUniversal, "UPDATE notas_debito SET Total = Total + " & Val(TxtValor.Text) & " WHERE IdNotaDebito = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
          AbrirRecorset rstUniversal, "UPDATE cuentas_cobrar SET Saldo = Saldo + " & Val(TxtValor.Text) & " WHERE IdCxC = " & rstCuentaCobrar!IdCxC, CnnPrincipal, adOpenDynamic, adLockOptimistic
          TxtCampos(3).Text = Val(TxtCampos(3).Text) + Val(TxtValor.Text)
          AccionTool 17
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
  Dim rstNotadebitoDet As New ADODB.Recordset
  rstNotadebitoDet.CursorLocation = adUseClient
  II = 1
  While II <= LstNotaDebitoDet.ListItems.Count
    If LstNotaDebitoDet.ListItems(II).Checked = True Then
      AbrirRecorset rstNotadebitoDet, "SELECT notas_debito_det.* FROM notas_debito_det WHERE IdNotaDebitoDet = " & LstNotaDebitoDet.SelectedItem, CnnPrincipal, adOpenDynamic, adLockOptimistic
        If rstNotadebitoDet.RecordCount > 0 Then
          AbrirRecorset rstUniversal, "UPDATE notas_debito SET Total = Total - " & rstNotadebitoDet!Valor & " WHERE IdNotaDebito = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
          AbrirRecorset rstUniversal, "UPDATE cuentas_cobrar SET Saldo = Saldo - " & rstNotadebitoDet!Valor & " WHERE IdCxC = " & rstNotadebitoDet!IdCxC, CnnPrincipal, adOpenDynamic, adLockOptimistic
          AbrirRecorset rstUniversal, "DELETE FROM notas_debito_det WHERE IdNotaDebitoDet=" & Val(LstNotaDebitoDet.SelectedItem), CnnPrincipal, adOpenDynamic, adLockOptimistic
          TxtCampos(3).Text = Val(TxtCampos(3).Text) - rstNotadebitoDet!Valor
        End If
      CerrarRecorset rstNotadebitoDet
      LstNotaDebitoDet.ListItems.Remove (II)
    Else
     II = II + 1
    End If
    
  Wend
End Sub

Private Sub CmdVerDetalles_Click()
  VerDetalle
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  LaTecla (KeyCode), ToolNotasDebito
End Sub

Private Sub Form_Load()
  IconosTool ToolNotasDebito, Principal.IgListTool
  rstNotasDebito.CursorLocation = adUseServer
  
  strSqlNotasDebito = "SELECT notas_debito.*, " & _
                "terceros.RazonSocial " & _
                "FROM notas_debito " & _
                "LEFT JOIN terceros ON notas_debito.IdTercero = terceros.IDTercero"
  AbrirRecorset rstNotasDebito, strSqlNotasDebito & " Order by IdNotaDebito Desc Limit 100", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Formatos rstNotasDebito
  Asignar rstNotasDebito
  Editando = False
End Sub

Private Sub Asignar(rstAsignar As ADODB.Recordset)
  For II = 0 To 5
    TxtCampos(II).Text = rstAsignar.Fields(II) & ""
  Next
  TxtNmTercero.Caption = rstAsignar!RazonSocial & ""
  ChkImpreso.Value = DevCheck(rstAsignar!Impreso)
  ChkAnulado.Value = DevCheck(rstAsignar!Anulado)
  If Val(rstAsignar!Impreso) = 1 Or Val(rstAsignar!Anulado) = 1 Then
    FraAgregar.Enabled = False
  Else
    FraAgregar.Enabled = True
  End If
  If LstNotaDebitoDet.Tag = "Llena" Then
    LstNotaDebitoDet.ListItems.Clear
    LstNotaDebitoDet.Tag = "Vacia"
  End If
  VerDetalle
End Sub
Private Sub Formatos(rstForma As ADODB.Recordset)
  For II = 0 To 5
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
            AbrirRecorset rstUniversal, "Update notas_debito SET IdTercero='" & TxtCampos(2).Text & "', Total= " & Val(TxtCampos(3).Text) & ", Comentarios = '" & TxtCampos(4).Text & "'  WHERE IdNotaDebito = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
            Editando = False
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
            'TxtCampos(0).Text = SacarConsecutivo("NotasDebito", CnnPrincipal)
            AbrirRecorset rstUniversal, "INSERT INTO notas_debito (Fecha, IdTercero, Total, Comentarios) " & _
                                        "VALUES ('" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', '" & TxtCampos(2).Text & "', " & Val(TxtCampos(3).Text) & ", '" & TxtCampos(4).Text & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
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
      'MsgBox "Esta opcion debe ser consultada con el desarrollador del sistema"
      If MsgBox("Desea anular esta nota debito?", vbQuestion + vbYesNo) = vbYes Then
        Dim rstActualizar As New ADODB.Recordset
        rstActualizar.CursorLocation = adUseClient
        AbrirRecorset rstUniversal, "SELECT notas_debito.* FROM notas_debito WHERE IdNotaDebito = " & Val(TxtCampos(0).Text) & " AND Impreso = 1 AND Anulado = 0", CnnPrincipal, adOpenDynamic, adLockOptimistic
        If rstUniversal.RecordCount > 0 Then
          AbrirRecorset rstUniversal, "SELECT notas_debito_det.* FROM notas_debito_det WHERE IdNotaDebito = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
          Do While rstUniversal.EOF = False
            AbrirRecorset rstActualizar, "UPDATE cuentas_cobrar SET Saldo = Saldo - " & rstUniversal!Valor & " WHERE IdCxC = " & rstUniversal!IdCxC, CnnPrincipal, adOpenDynamic, adLockOptimistic
            rstUniversal.MoveNext
          Loop
          CerrarRecorset rstUniversal
          AbrirRecorset rstActualizar, "UPDATE notas_debito_det SET Valor = 0 WHERE IdNotaDebito = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
          AbrirRecorset rstActualizar, "UPDATE notas_debito SET Total = 0, Anulado = 1 WHERE IdNotaDebito = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
          MsgBox "Nota debito anulado con exito", vbInformation
          AccionTool 17
          Asignar rstNotasDebito
        Else
          MsgBox "La nota debito debe estar impreso y sin anular para poder ser anulado", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If
    Case 7 'Cancelar
      If Editando = True Then
        AccionTool 4
      Else
        Asignar rstNotasDebito
        Bloquear
      End If
    Case 9  'Buscar
      If Principal.ToolConsultas1.AbrirDevDatos("Numero de nota debito", "Digite el numero del nota debito que desea buscar", 3, 0) = True Then
        AbrirRecorset rstUniversal, strSqlNotasDebito & " WHERE numeroNotaDebito = " & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          Asignar rstUniversal
        Else
          MsgBox "No se encontraron notas debito con este numero", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If
    Case 11 'Primero
      UPrimero rstNotasDebito
      Asignar rstNotasDebito
    Case 12 'Anterior
      UAnterior rstNotasDebito
      Asignar rstNotasDebito
    Case 13 'Siguiente
      USiguiente rstNotasDebito
      Asignar rstNotasDebito
    Case 14 'Ultimo
      UUltimo rstNotasDebito
      Asignar rstNotasDebito
    Case 16 'Cerrar
      Set rstNotasDebito = Nothing
      Unload Me
    Case 17 'Actualizar
      rstNotasDebito.Requery
      Formatos rstNotasDebito
    Case 18 'Imprimir
      AbrirRecorset rstUniversal, "SELECT notas_debito.* FROM notas_debito WHERE IdNotaDebito = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.RecordCount > 0 Then
        If Val(rstUniversal!Impreso) = 0 Then
          FufuLo = SacarConsecutivo("NotasDebito", CnnPrincipal)
          AbrirRecorset rstUniversal, "UPDATE notas_debito SET Impreso = 1, numeroNotaDebito = " & FufuLo & " WHERE IdNotaDebito=" & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
          Mostrar_Reporte CnnPrincipal, 55, "SELECT sql_ic_imprimir_nota_debito.* FROM sql_ic_imprimir_nota_debito WHERE IdNotaDebito = " & Val(TxtCampos(0).Text), "Imprimir nota debito", 2
          AccionTool 17
          Asignar rstNotasDebito
        Else
          Mostrar_Reporte CnnPrincipal, 55, "SELECT sql_ic_imprimir_nota_debito.* FROM sql_ic_imprimir_nota_debito WHERE IdNotaDebito = " & Val(TxtCampos(0).Text), "Imprimir nota debito", 2
        End If
      End If
      CerrarRecorset rstUniversal
    Case 19
    Case 20
  End Select
End Sub

Private Sub Desbloquear()
  BotTool 3, 17, ToolNotasDebito, True
  FraDatos.Enabled = True
  FraAgregar.Enabled = False
  CmdVerDetalles.Enabled = False
End Sub

Private Sub Bloquear()
  BotTool 3, 17, ToolNotasDebito, False
  FraDatos.Enabled = False
  CmdVerDetalles.Enabled = True
End Sub

Private Sub Limpiar()
  For II = 0 To 5
    TxtCampos(II).Text = ""
  Next
  TxtNmTercero.Caption = ""
  LstNotaDebitoDet.ListItems.Clear
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

Private Sub ToolNotasDebito_ButtonClick(ByVal Button As MSComctlLib.Button)
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
  If rstUniversal.RecordCount < 0 Then
    TxtCuentaCobrar.Text = ""
    TxtCuentaCobrar.SetFocus
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
  Dim rstNotasDebitoDetalle As New ADODB.Recordset
  rstNotasDebitoDetalle.CursorLocation = adUseClient
  AbrirRecorset rstNotasDebitoDetalle, "SELECT notas_debito_det.*, NroDocumento, FechaDoc " & _
                          "FROM notas_debito_det " & _
                          "LEFT JOIN cuentas_cobrar ON notas_debito_det.IdCxC = cuentas_cobrar.IdCxC " & _
                          "WHERE IdNotaDebito = " & Val(TxtCampos(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
  LstNotaDebitoDet.ListItems.Clear
  If rstNotasDebitoDetalle.RecordCount > 0 Then
    Do While rstNotasDebitoDetalle.EOF = False
      Set Item = LstNotaDebitoDet.ListItems.Add(, , rstNotasDebitoDetalle!IdNotaDebitoDet)
      Item.SubItems(1) = rstNotasDebitoDetalle!IdCxC
      Item.SubItems(2) = rstNotasDebitoDetalle!NroDocumento & ""
      Item.SubItems(3) = rstNotasDebitoDetalle!FechaDoc & ""
      Item.SubItems(4) = rstNotasDebitoDetalle!Valor
      rstNotasDebitoDetalle.MoveNext
    Loop
  End If
  LstNotaDebitoDet.Tag = "Llena"
End Sub




