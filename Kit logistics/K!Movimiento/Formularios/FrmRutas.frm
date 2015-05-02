VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRutas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rutas..."
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9225
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraCiudades 
      Height          =   5175
      Left            =   4560
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton CmdQuitar 
         Caption         =   "&Quitar"
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton CmdGuardarCambios 
         Caption         =   "&Guardar cambios"
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   4800
         Width           =   2055
      End
      Begin VB.CommandButton CmdDejarIgual 
         Caption         =   "&No cambiar"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   4800
         Width           =   2175
      End
      Begin VB.CommandButton CmdBajar 
         Caption         =   "&Bajar"
         Height          =   255
         Left            =   3480
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton CmdSubir 
         Caption         =   "&Subir"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   840
         Width           =   975
      End
      Begin MSComctlLib.ListView LstCiudades 
         Height          =   3495
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   6165
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ciudad"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.TextBox TxtIdCiudad 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   0
         Tag             =   "1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label LblConsulta 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   12
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "ID Ciudad:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "Editar >>"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Tag             =   "0"
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton CmdVer 
      Caption         =   "Ver"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4335
      Begin VB.TextBox TxtNmRuta 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox TxtIdRuta 
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   165
      End
   End
   Begin MSDataGridLib.DataGrid GrillaRutas 
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7011
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "IdCIudad"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "NmCiudad"
         Caption         =   "Ciudad"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Orden"
         Caption         =   "Orden"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2805.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   569.764
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolRutas 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   9225
      _ExtentX        =   16272
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
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Reporte de rutas"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmRutas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstRutas As New ADODB.Recordset
Dim rstTem As New ADODB.Recordset
Dim Editando As Boolean, GrillaLlena As Boolean, ElIndex As Integer, Itemtem As ListItem

Private Sub CmdAgregar_Click()
  If Val(TxtIdCiudad) <> 0 Then
    AgregarCiudad
  End If
End Sub

Private Sub CmdBajar_Click()
  If LstCiudades.ListItems.Count > 1 Then
    Set Itemtem = LstCiudades.SelectedItem
    If Itemtem.Index < LstCiudades.ListItems.Count Then
      ElIndex = Itemtem.Index
      FufuLo = Itemtem
      FufuSt = Itemtem.SubItems(1)
      LstCiudades.ListItems.Remove (Itemtem.Index)
      Set Item = LstCiudades.ListItems.Add(ElIndex + 1, , FufuLo)
      Item.SubItems(1) = FufuSt
      Set LstCiudades.SelectedItem = LstCiudades.ListItems(ElIndex + 1)
    End If
  End If
End Sub

Private Sub CmdDejarIgual_Click()
  If MsgBox("¿Esta seguro de que no desea guardar cambios de las acciones realizadas?", vbQuestion + vbYesNo, "No se guardaran los cambios") = vbYes Then
    LstCiudades.ListItems.Clear
    FraCiudades.Visible = False
    CmdVer.Enabled = True
    CmdEditar.Enabled = True
    ToolRutas.Enabled = True
  End If
End Sub

Private Sub CmdEditar_Click()
  Bloquear
  FraCiudades.Visible = True
  TxtIdCiudad.SetFocus
  CmdEditar.Tag = 1
  ToolRutas.Enabled = False
  CmdEditar.Enabled = False
  CmdVer.Enabled = False
  LstCiudades.ListItems.Clear
  AbrirRecorset rstUniversal, "SELECT rutas_ciudades.IdRuta, rutas_ciudades.IdCiudad, ciudades.NmCiudad,  rutas_ciudades.Orden From  rutas_ciudades INNER JOIN ciudades ON (rutas_ciudades.IdCiudad = ciudades.IdCiudad) Where rutas_ciudades.IdRuta=" & Val(TxtIdRuta) & " Order by Orden", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Do While rstUniversal.EOF = False
    Set Item = LstCiudades.ListItems.Add(, , rstUniversal!IdCiudad)
      Item.SubItems(1) = rstUniversal!NmCiudad
      rstUniversal.MoveNext
    Loop
  CerrarRecorset rstUniversal
End Sub
Sub LimpiarGrilla()
  If GrillaLlena = True Then
    CerrarRecorset rstTem
    GrillaLlena = False
  End If
End Sub

Private Sub CmdGuardarCambios_Click()
  If MsgBox("¿Esta seguro de que desea realizar estos cambios?", vbQuestion + vbYesNo, "Realizar cambios") = vbYes Then
    AbrirRecorset rstUniversal, "Delete from rutas_ciudades where idruta=" & Val(TxtIdRuta), CnnPrincipal, adOpenDynamic, adLockOptimistic
    For II = 1 To LstCiudades.ListItems.Count
      AbrirRecorset rstUniversal, "insert into rutas_ciudades (IdRuta, IdCiudad, orden) values (" & Val(TxtIdRuta.Text) & ", " & LstCiudades.ListItems(II) & ", " & II & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
    Next
    MsgBox "Lista cambiada con exito", vbInformation
    LstCiudades.ListItems.Clear
    FraCiudades.Visible = False
    CmdVer.Enabled = True
    CmdEditar.Enabled = True
    ToolRutas.Enabled = True
  End If
End Sub

Private Sub CmdQuitar_Click()
II = 1
While II <= LstCiudades.ListItems.Count
  If LstCiudades.ListItems(II).Checked = True Then
    LstCiudades.ListItems.Remove (II)
  Else
   II = II + 1
  End If
Wend
End Sub

Private Sub CmdSubir_Click()
  If LstCiudades.ListItems.Count > 1 Then
    Set Itemtem = LstCiudades.SelectedItem
    If Itemtem.Index > 1 Then
      ElIndex = Itemtem.Index
      FufuLo = Itemtem
      FufuSt = Itemtem.SubItems(1)
      LstCiudades.ListItems.Remove (Itemtem.Index)
      Set Item = LstCiudades.ListItems.Add(ElIndex - 1, , FufuLo)
      Item.SubItems(1) = FufuSt
      Set LstCiudades.SelectedItem = LstCiudades.ListItems(ElIndex - 1)
    End If
  End If
End Sub

Private Sub CmdVer_Click()
  AbrirRecorset rstTem, "SELECT rutas_ciudades.IdRuta, rutas_ciudades.IdCiudad, ciudades.NmCiudad, rutas_ciudades.Orden from rutas_ciudades  INNER JOIN ciudades ON (rutas_ciudades.IdCiudad = ciudades.IdCiudad) where rutas_ciudades.IdRuta=" & TxtIdRuta & " order by Orden", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Set GrillaRutas.DataSource = rstTem
  GrillaLlena = True
  
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  LaTecla (KeyCode), ToolRutas
End Sub
Private Sub Form_Load()
  IconosTool ToolRutas, Principal.IgListTool
  rstRutas.CursorLocation = adUseServer
  rstTem.CursorLocation = adUseClient
  AbrirRecorset rstRutas, "SELECT*From Rutas", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Asignar rstRutas
End Sub
Private Sub Asignar(rstAsignar As ADODB.Recordset)
  TxtIdRuta.Text = rstAsignar!IdRuta
  TxtNmRuta.Text = rstAsignar!NmRuta & ""
End Sub
Private Sub limpiar()
  TxtIdRuta.Text = ""
  TxtNmRuta.Text = ""
End Sub
Private Sub Desbloquear()
  FraDatos.Enabled = True
  BotTool 3, 17, ToolRutas, True
  CmdEditar.Enabled = False
  CmdVer.Enabled = False
End Sub
Private Sub Bloquear()
  FraDatos.Enabled = False
  BotTool 3, 17, ToolRutas, False
  CmdEditar.Enabled = True
  CmdVer.Enabled = True
End Sub
Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
      Desbloquear
      limpiar
      TxtNmRuta.SetFocus
      LimpiarGrilla
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            Bloquear
            AbrirRecorset rstUniversal, "Update Rutas set NmRuta='" & TxtNmRuta & "' where IdRuta=" & Val(TxtIdRuta), CnnPrincipal, adOpenDynamic, adLockOptimistic
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          AbrirRecorset rstUniversal, "INSERT INTO Rutas (NmRuta) VALUES ('" & TxtNmRuta & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
          Asignar rstRutas
          Bloquear
        End If
      End If
    Case 5  'Editar
      Editando = True
      Desbloquear
    Case 6 'Eliminar
      If MsgBox("¿Esta seguro de eliminar esta ruta?" & Chr(13) & "Recuerde que al eliminar una ruta se elimina la informacio que depende de ella", vbYesNo + vbQuestion) = vbYes Then
        AbrirRecorset rstUniversal, "Delete from rutas where Idruta=" & Val(TxtIdRuta), CnnPrincipal, adOpenDynamic, adLockOptimistic
        AbrirRecorset rstUniversal, "Delete from rutas_ciudades where idruta=" & Val(TxtIdRuta), CnnPrincipal, adOpenDynamic, adLockOptimistic
        AccionTool 17
        Asignar rstRutas
      End If
    Case 7 'Cancelar
      If Editando = True Then
        AccionTool 4
      Else
        Asignar rstRutas
        Bloquear
      End If
    Case 9  'Buscar
      If Principal.ToolConsultas1.AbrirDevConsultaCO(2, Coperaciones, CnnPrincipal) = True Then
        AbrirRecorset rstUniversal, "Select*from Rutas where IdRuta=" & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          Asignar rstUniversal
        Else
          MsgBox "No se encontraron rutas con este codigo", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If
      
    Case 11 'Primero
      UPrimero rstRutas
      Asignar rstRutas
      LimpiarGrilla
    Case 12 'Anterior
      UAnterior rstRutas
      Asignar rstRutas
      LimpiarGrilla
    Case 13 'Siguiente
      USiguiente rstRutas
      Asignar rstRutas
      LimpiarGrilla
    Case 14 'Ultimo
      UUltimo rstRutas
      Asignar rstRutas
      LimpiarGrilla
    Case 16 'Cerrar
      CerrarRecorset rstRutas
      CerrarRecorset rstUniversalAux
      Unload Me
    Case 17 'Actualizar
      rstRutas.Requery
    Case 18 'Imprimir
  End Select
End Sub
Private Sub TxtIdCiudad_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Principal.ToolConsultas1.AbrirDevConsulta 1, CnnPrincipal
    TxtIdCiudad.Text = Principal.ToolConsultas1.DatLo
  End If
End Sub
Private Sub TxtIdCiudad_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  ValidarEntrada TxtIdCiudad, KeyAscii, 1
End Sub
Private Sub TxtIdCiudad_Validate(Cancel As Boolean)
    If Val(TxtIdCiudad.Text) <> 0 Then
      AbrirRecorset rstUniversal, "SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & TxtIdCiudad, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        LblConsulta(1) = rstUniversal!NmCiudad & ""
      Else
        LblConsulta(1) = "": TxtIdCiudad.Text = ""
      End If
      CerrarRecorset rstUniversal
    End If
End Sub
Private Sub ToolRutas_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub
Sub AgregarCiudad()
  Set Item = LstCiudades.FindItem(TxtIdCiudad)
  If Item Is Nothing Then
    Set Item = LstCiudades.ListItems.Add(, , Val(TxtIdCiudad))
    Item.SubItems(1) = LblConsulta(1)
    TxtIdCiudad.Text = ""
    LblConsulta(1).Caption = ""
    TxtIdCiudad.SetFocus
  Else
    MsgBox "Esta ciudad ya se encuentra en la lista", vbCritical
  End If
End Sub

Function Validacion() As Boolean
  If TxtNmRuta.Text <> "" Then
    Validacion = True
  Else
    MsgTit "La Ruta debe tener un nombre": TxtNmRuta.SetFocus: Validacion = False
  End If

End Function

