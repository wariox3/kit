VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRelEntrega 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relaciones de entrega de documentos..."
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8820
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCargarGuias 
      Caption         =   "Cargar guias"
      Height          =   255
      Left            =   6720
      TabIndex        =   11
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   8535
      Begin VB.TextBox TxtCampo 
         Height          =   1125
         Index           =   3
         Left            =   1080
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   960
         Width           =   7335
      End
      Begin VB.TextBox TxtCampo 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   1080
         MaxLength       =   11
         TabIndex        =   0
         ToolTipText     =   "Aqui se debe ingresar el tercero"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox TxtNmTercero 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2880
         MaxLength       =   60
         TabIndex        =   8
         ToolTipText     =   "Aqui se debe ingresar el tercero"
         Top             =   600
         Width           =   5535
      End
      Begin VB.TextBox TxtCampo 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtCampo 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Comentario:"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   10
         Top             =   960
         Width           =   840
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Tercero:"
         Height          =   195
         Index           =   33
         Left            =   300
         TabIndex        =   9
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         Height          =   195
         Left            =   690
         TabIndex        =   6
         Top             =   240
         Width           =   210
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
   End
   Begin MSComctlLib.Toolbar ToolRelEntrega 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   1005
      ButtonWidth     =   847
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
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
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Acc"
            Style           =   5
            Object.Width           =   1e-4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmRelEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstRelEntrega As New ADODB.Recordset
Dim Editando As Boolean
Dim strSqlRelEntegaDoc As String

Private Sub CmdCargarGuias_Click()
  FufuSt = "N"
  FufuLo = TxtCampo(0).Text
  FrmLlenarRelEntegaDoc.Show 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  LaTecla (KeyCode), ToolRelEntrega
End Sub
Private Sub Form_Load()
  IconosTool ToolRelEntrega, Principal.IgListTool
  rstRelEntrega.CursorLocation = adUseServer
  strSqlRelEntegaDoc = "SELECT relentregadoc.*, " & _
                        "terceros.RazonSocial " & _
                        "FROM relentregadoc " & _
                        "LEFT JOIN terceros ON relentregadoc.IdTercero = terceros.IDTercero "
  AbrirRecorset rstRelEntrega, strSqlRelEntegaDoc, CnnPrincipal, adOpenDynamic, adLockOptimistic
  Asignar rstRelEntrega
End Sub

Private Sub Asignar(rstAsignar As ADODB.Recordset)
  For II = 0 To 3
    TxtCampo(II) = rstAsignar.Fields(II) & ""
  Next
  TxtNmTercero.Text = rstAsignar!RazonSocial
End Sub

Private Sub limpiar()
  For II = 0 To 3
    TxtCampo(II).Text = ""
  Next
  TxtNmTercero.Text = ""
End Sub
Private Sub Desbloquear()
  FraDatos.Enabled = True
  BotTool 3, 17, ToolRelEntrega, True
End Sub
Private Sub Bloquear()
  FraDatos.Enabled = False
  BotTool 3, 17, ToolRelEntrega, False
End Sub
Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
      Desbloquear
      limpiar
      TxtCampo(1).Text = Format(Date, "dd/mm/yy")
      TxtCampo(2).SetFocus
      Editando = False
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            Bloquear
            AbrirRecorset rstUniversal, "Update relentregadoc set Comentarios='" & TxtCampo(2).Text & "', Comentarios='" & TxtCampo(3).Text & "' where IDRel=" & Val(TxtCampo(0).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          AbrirRecorset rstUniversal, "INSERT INTO relentregadoc (IdTercero, Fecha, Comentarios) VALUES ('" & TxtCampo(2).Text & "', '" & Format(Date, "yy/mm/dd") & "', '" & TxtCampo(3).Text & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
          Bloquear
          AccionTool 17
          AccionTool 14
        End If
      End If
    Case 5  'Editar
      Editando = True
      Desbloquear
    Case 6 'Eliminar
    Case 7 'Cancelar
      If Editando = True Then
        AccionTool 4
      Else
        Asignar rstRelEntrega
        Bloquear
      End If
    Case 9  'Buscar
      If Principal.ToolConsultas1.AbrirDevDatos("Numero relacion", "Digite el numero de la relacion", 3, 0) = True Then
        AbrirRecorset rstUniversal, strSqlRelEntegaDoc & " WHERE IDRel=" & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          Asignar rstUniversal
        Else
          MsgBox "No se encontraron relaciones con este numero", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If
    Case 11 'Primero
      UPrimero rstRelEntrega
      Asignar rstRelEntrega
    Case 12 'Anterior
      UAnterior rstRelEntrega
      Asignar rstRelEntrega
    Case 13 'Siguiente
      USiguiente rstRelEntrega
      Asignar rstRelEntrega
    Case 14 'Ultimo
      UUltimo rstRelEntrega
      Asignar rstRelEntrega
    Case 16 'Cerrar
      CerrarRecorset rstRelEntrega
      FufuSt = TxtCampo(0)
      Unload Me
    Case 17 'Actualizar
      rstRelEntrega.Requery
    Case 18 'Imprimir
      Mostrar_Reporte CnnPrincipal, 39, "SELECT sql_im_rel_guias_cumplidos.* FROM sql_im_rel_guias_cumplidos WHERE IdRelEntrega=" & Val(TxtCampo(0).Text), "", 2
    Case 19
      'If TxtCampo(9).Text <> "" Then TxtNmCliente.Text = DevResBus("SELECT Id, NmNegociacion From Negociaciones where Id=" & Val(TxtCampo(9)), "NmNegociacion", CnnPrincipal)
  End Select
End Sub
Function Validacion() As Boolean
  Validacion = False
  If TxtCampo(2).Text <> "" Then
    Validacion = True
  Else
    MsgBox "La relacion debe tener un tercero", vbCritical: TxtCampo(2).SetFocus
  End If
End Function

Private Sub ToolRelEntrega_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub

Private Sub TxtCampo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Select Case Index
      Case 2
        Principal.ToolConsultas1.AbrirDevConsulta 7, CnnPrincipal
        TxtCampo(2).Text = Principal.ToolConsultas1.DatSt
    End Select
  End If
End Sub

Private Sub TxtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
    Case 2
      ValidarEntrada TxtCampo(Index), KeyAscii, 1
    Case 3
      If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys vbTab
      End If
  End Select
  If KeyAscii = 13 Then
      SendKeys vbTab
  End If
End Sub

Private Sub TxtCampo_LostFocus(Index As Integer)
  Select Case Index
    Case 2
      If TxtCampo(2).Text <> "" Then
        AbrirRecorset rstUniversal, "Select IdTercero, RazonSocial from Terceros where IdTercero='" & TxtCampo(2) & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
          If rstUniversal.EOF = False Then
            TxtNmTercero.Text = rstUniversal.Fields("RazonSocial") & ""
            CerrarRecorset rstUniversal
          Else
            TxtCampo(2).Text = ""
            TxtNmTercero.Text = ""
          End If
        CerrarRecorset rstUniversal
      End If
  End Select
End Sub
