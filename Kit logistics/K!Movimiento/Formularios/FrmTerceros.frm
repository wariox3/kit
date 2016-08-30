VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTerceros 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Terceros..."
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdInactivar 
      Caption         =   "Inactivar"
      Height          =   255
      Left            =   3840
      TabIndex        =   30
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton CmdCambiarNit 
      Caption         =   "Pasar movimientos a otro nit"
      Height          =   255
      Left            =   6000
      TabIndex        =   28
      Top             =   720
      Width           =   2895
   End
   Begin TabDlg.SSTab SSTTerceros 
      Height          =   4335
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos"
      TabPicture(0)   =   "FrmTerceros.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Comercial"
      TabPicture(1)   =   "FrmTerceros.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraNegociaciones"
      Tab(1).Control(1)=   "FraDatosComerciales"
      Tab(1).ControlCount=   2
      Begin VB.Frame FraNegociaciones 
         Enabled         =   0   'False
         Height          =   1935
         Left            =   -74880
         TabIndex        =   39
         Top             =   2280
         Width           =   8535
         Begin VB.CheckBox ChkCorriente 
            Caption         =   "Corriente"
            Height          =   255
            Left            =   7200
            TabIndex        =   54
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox ChkDestino 
            Caption         =   "Destino"
            Height          =   255
            Left            =   7200
            TabIndex        =   53
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox ChkContado 
            Caption         =   "Contado"
            Height          =   255
            Left            =   7200
            TabIndex        =   52
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdQuitar 
            Caption         =   "Quitar"
            Height          =   255
            Left            =   4440
            TabIndex        =   42
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton CmdAgregarNegociacion 
            Caption         =   "Agregar"
            Height          =   255
            Left            =   4440
            TabIndex        =   41
            Top             =   240
            Width           =   975
         End
         Begin MSComctlLib.ListView LstNegociaciones 
            Height          =   1575
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   2778
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
               Text            =   "Id"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Negociacion"
               Object.Width           =   5292
            EndProperty
         End
      End
      Begin VB.Frame FraDatosComerciales 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   -74880
         TabIndex        =   33
         Top             =   360
         Width           =   8535
         Begin VB.TextBox TxtCampo 
            Height          =   285
            Index           =   16
            Left            =   4920
            TabIndex        =   47
            Top             =   600
            Width           =   3495
         End
         Begin VB.TextBox TxtNmCentroCostos 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3360
            TabIndex        =   49
            Top             =   1320
            Width           =   5055
         End
         Begin VB.TextBox TxtCampo 
            Height          =   285
            Index           =   15
            Left            =   1440
            TabIndex        =   46
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox TxtNmAsesor 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3360
            TabIndex        =   48
            Top             =   960
            Width           =   5055
         End
         Begin VB.TextBox TxtCampo 
            Height          =   285
            Index           =   14
            Left            =   1440
            TabIndex        =   45
            Top             =   960
            Width           =   1815
         End
         Begin VB.ComboBox CboFormaPago 
            Height          =   315
            ItemData        =   "FrmTerceros.frx":0038
            Left            =   1440
            List            =   "FrmTerceros.frx":0042
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox TxtCampo 
            Height          =   285
            Index           =   12
            Left            =   1440
            TabIndex        =   34
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Condicion comercial:"
            Height          =   195
            Left            =   3360
            TabIndex        =   51
            Top             =   600
            Width           =   1470
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Centro costos:"
            Height          =   195
            Left            =   345
            TabIndex        =   50
            Top             =   1320
            Width           =   1020
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Asesor:"
            Height          =   195
            Left            =   810
            TabIndex        =   44
            Top             =   960
            Width           =   525
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Plazo:"
            Height          =   195
            Left            =   900
            TabIndex        =   37
            Top             =   600
            Width           =   435
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Forma Pago:"
            Height          =   195
            Left            =   435
            TabIndex        =   36
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.Frame FraDatos 
         Enabled         =   0   'False
         Height          =   3855
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   8535
         Begin VB.TextBox TxtCampo 
            Height          =   285
            Index           =   13
            Left            =   4440
            TabIndex        =   0
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton CmdCrearNegociacion 
            Caption         =   "Crear"
            Height          =   255
            Left            =   7800
            TabIndex        =   38
            Top             =   3480
            Width           =   615
         End
         Begin VB.TextBox TxtCampo 
            Height          =   285
            Index           =   11
            Left            =   1080
            TabIndex        =   9
            Top             =   2760
            Width           =   1815
         End
         Begin VB.TextBox TxtCampo 
            Height          =   285
            Index           =   10
            Left            =   3480
            TabIndex        =   8
            Top             =   2400
            Width           =   4935
         End
         Begin VB.TextBox TxtNmCliente 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   17
            Top             =   3480
            Width           =   5895
         End
         Begin VB.TextBox TxtNmCiudad 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   16
            Top             =   3120
            Width           =   6615
         End
         Begin VB.TextBox TxtCampo 
            Height          =   285
            Index           =   9
            Left            =   1080
            TabIndex        =   11
            Top             =   3480
            Width           =   615
         End
         Begin VB.TextBox TxtCampo 
            Height          =   285
            Index           =   8
            Left            =   1080
            TabIndex        =   10
            Top             =   3120
            Width           =   615
         End
         Begin VB.TextBox TxtCampo 
            Height          =   285
            Index           =   7
            Left            =   1080
            MaxLength       =   7
            TabIndex        =   7
            Top             =   2400
            Width           =   1815
         End
         Begin VB.TextBox TxtCampo 
            Height          =   285
            Index           =   1
            Left            =   5760
            MaxLength       =   1
            TabIndex        =   1
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox TxtCampo 
            Height          =   285
            Index           =   6
            Left            =   1080
            TabIndex        =   6
            Top             =   2040
            Width           =   7335
         End
         Begin VB.TextBox TxtCampo 
            Height          =   285
            Index           =   5
            Left            =   1080
            TabIndex        =   4
            Top             =   1320
            Width           =   5175
         End
         Begin VB.TextBox TxtCampo 
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   3
            Top             =   960
            Width           =   5175
         End
         Begin VB.TextBox TxtCampo 
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   2
            Top             =   600
            Width           =   5175
         End
         Begin VB.TextBox TxtCampo 
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   5
            Top             =   1680
            Width           =   7335
         End
         Begin VB.TextBox TxtCampo 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   15
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Digito verificacion:"
            Height          =   195
            Left            =   3000
            TabIndex        =   43
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Celular:"
            Height          =   195
            Left            =   360
            TabIndex        =   32
            Top             =   2760
            Width           =   525
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Email:"
            Height          =   195
            Left            =   3000
            TabIndex        =   31
            Top             =   2400
            Width           =   420
         End
         Begin VB.Label LblEstado 
            Alignment       =   2  'Center
            Caption         =   "Inactivo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6360
            TabIndex        =   29
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Negociacion:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   3480
            Width           =   945
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Ciudad:"
            Height          =   195
            Left            =   360
            TabIndex        =   26
            Top             =   3120
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Telefono:"
            Height          =   195
            Left            =   225
            TabIndex        =   25
            Top             =   2400
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Direccion:"
            Height          =   195
            Left            =   180
            TabIndex        =   24
            Top             =   2040
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Apellido2:"
            Height          =   195
            Left            =   210
            TabIndex        =   23
            Top             =   1320
            Width           =   690
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Apellido1:"
            Height          =   195
            Left            =   210
            TabIndex        =   22
            Top             =   960
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   600
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Extendido:"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1680
            Width           =   750
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   5280
            TabIndex        =   19
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ID:"
            Height          =   195
            Left            =   690
            TabIndex        =   18
            Top             =   240
            Width           =   210
         End
      End
   End
   Begin MSComctlLib.Toolbar ToolTerceros 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
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
Attribute VB_Name = "FrmTerceros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstTerceros As New ADODB.Recordset
Dim rstAct As New ADODB.Recordset
Dim Editando As Boolean
Dim strTerceros As String

Private Sub CmdAgregarNegociacion_Click()
  Principal.ToolConsultas1.AbrirDevConsulta 2, CnnPrincipal
  If Principal.ToolConsultas1.DatLo <> 0 Then
    AbrirRecorset rstUniversal, "INSERT INTO negociaciones_terceros (IdTercero, IdNegociacion) VALUES ('" & TxtCampo(0).Text & "', " & Principal.ToolConsultas1.DatLo & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
    VerNegociaciones
  End If
End Sub

Private Sub CmdCambiarNit_Click()
  Dim NuevoNit As String, ViejoNit As String
  If CpPermisoEspecial(5, CodUsuarioActivo, CnnPrincipal) = True Then
    ViejoNit = TxtCampo(0).Text
    NuevoNit = InputBox("Digite el nit para el cual se le van a asignar los movimientos de este nit (Facturas-Guias)", "Digite el nuevo nit")
    If NuevoNit <> "" And NuevoNit <> ViejoNit Then
      AbrirRecorset rstUniversal, "Select IdTercero from terceros where IdTercero='" & NuevoNit & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.EOF = False Then
        CerrarRecorset rstUniversal
        AbrirRecorset rstUniversal, "Update Facturas set IdCliente='" & NuevoNit & "' where IdCliente='" & ViejoNit & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
        AbrirRecorset rstUniversal, "Update Guias set Cuenta='" & NuevoNit & "' where Cuenta='" & ViejoNit & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
        AbrirRecorset rstUniversal, "Update cuentas_cobrar set IdTercero='" & NuevoNit & "' where IdTercero='" & ViejoNit & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
        AbrirRecorset rstUniversal, "Update facturas_venta set IdTercero='" & NuevoNit & "' where IdTercero='" & ViejoNit & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
        
        MsgBox "La informacion del nit " & ViejoNit & " se paso exitosamente al nit " & NuevoNit, vbInformation
      Else
        MsgBox "El nit" & NuevoNit & " no existe para pasarle los movimientos de este nit", vbCritical
      End If
      CerrarRecorset rstUniversal
    End If
  End If
End Sub

Private Sub CmdCrearNegociacion_Click()
  FrmClientesNegociacion.Show 1
  If Editando = True Then
    If Val(TxtCampo(9)) = 0 Then
      TxtCampo(9).Text = FufuLo
    End If
    TxtCampo(9).SetFocus
  End If
End Sub

Private Sub CmdInactivar_Click()
  If CpPermisoEspecial(17, CodUsuarioActivo, CnnPrincipal) = True Then
    rstAct.CursorLocation = adUseClient
    AbrirRecorset rstUniversal, "Select IDTercero, Inactivo from terceros where IDTercero='" & Val(TxtCampo(0).Text) & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstUniversal.EOF = False Then
      If Val(rstUniversal.Fields("Inactivo")) = 0 Then
        AbrirRecorset rstAct, "Update terceros set Inactivo=1 where IDTercero='" & Val(TxtCampo(0).Text) & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
        LblEstado.Caption = "Inactivo"
        CmdInactivar.Caption = "Activar"
      Else
        AbrirRecorset rstAct, "Update terceros set Inactivo=0 where IDTercero='" & Val(TxtCampo(0).Text) & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
        LblEstado.Caption = "Activo"
        CmdInactivar.Caption = "Inactivar"
      End If
    End If
    CerrarRecorset rstUniversal
  Else
    MsgBox "No tiene permisos para esta opcion", vbCritical
  End If
End Sub

Private Sub CmdQuitar_Click()
  II = 1
  While II <= LstNegociaciones.ListItems.Count
    If LstNegociaciones.ListItems(II).Checked = True Then
      AbrirRecorset rstUniversal, "DELETE FROM negociaciones_terceros WHERE IdTercero='" & TxtCampo(0).Text & "' AND IdNegociacion=" & LstNegociaciones.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
      LstNegociaciones.ListItems.Remove (II)
    Else
     II = II + 1
    End If
  Wend
  VerNegociaciones
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  LaTecla (KeyCode), ToolTerceros
End Sub
Private Sub Form_Load()
  IconosTool ToolTerceros, Principal.IgListTool
  rstTerceros.CursorLocation = adUseServer
  strTerceros = "SELECT terceros.*, " & _
                             "ciudades.NmCiudad, " & _
                             "negociaciones.NmNegociacion, " & _
                             "asesores.NmAsesor, " & _
                             "centros_costos.NmCentroCostos " & _
                             "FROM terceros " & _
                             "LEFT JOIN ciudades ON terceros.IdCiudad = ciudades.IdCiudad " & _
                             "LEFT JOIN asesores ON terceros.IdAsesor = asesores.IdAsesor " & _
                             "LEFT JOIN centros_costos ON terceros.IdCentroCostos = centros_costos.IdCentroCostos " & _
                             "LEFT JOIN negociaciones ON terceros.IdCliente = negociaciones.Id"
  AbrirRecorset rstTerceros, strTerceros, CnnPrincipal, adOpenDynamic, adLockOptimistic
  Asignar rstTerceros
End Sub

Private Sub Asignar(rstAsignar As ADODB.Recordset)
  For II = 0 To 16
    TxtCampo(II) = rstAsignar.Fields(II) & ""
  Next
  TxtNmCiudad.Text = rstAsignar!NmCiudad & ""
  TxtNmCliente.Text = rstAsignar!NmNegociacion & ""
  TxtNmAsesor.Text = rstAsignar!NmAsesor & ""
  TxtNmCentroCostos.Text = rstAsignar!NmCentroCostos & ""
  CboFormaPago.ListIndex = rstAsignar!IdFormaPago - 1
  If Val(rstAsignar.Fields("Inactivo")) = 0 Then
    LblEstado.Caption = "Activo"
    CmdInactivar.Caption = "Inactivar"
  Else
    LblEstado.Caption = "Inactivo"
    CmdInactivar.Caption = "Activar"
  End If
  ChkContado.value = DevCheck(rstAsignar!ManejaCobroContado)
  ChkDestino.value = DevCheck(rstAsignar!ManejaCobroDestino)
  ChkCorriente.value = DevCheck(rstAsignar!ManejaCobroCorriente)
  VerNegociaciones
End Sub

Private Sub limpiar()
  For II = 0 To 16
    TxtCampo(II).Text = ""
  Next
  TxtNmCiudad.Text = ""
  TxtNmCliente.Text = ""
  TxtNmAsesor.Text = ""
  TxtNmCentroCostos.Text = ""
  LstNegociaciones.ListItems.Clear
End Sub
Private Sub Desbloquear()
  FraDatos.Enabled = True
  FraDatosComerciales.Enabled = True
  FraNegociaciones.Enabled = True
  BotTool 3, 17, ToolTerceros, True
  CmdInactivar.Enabled = False
End Sub
Private Sub Bloquear()
  FraDatos.Enabled = False
  FraDatosComerciales.Enabled = False
  FraNegociaciones.Enabled = False
  BotTool 3, 17, ToolTerceros, False
  CmdInactivar.Enabled = True
End Sub
Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
      If Principal.ToolConsultas1.AbrirDevDatos("Digite ID", "Digite la identificacion del destinatario Nit/CC/CE", 2, 0) = True Then
        FufuSt = Principal.ToolConsultas1.DatSt
        If ExRecorset("Select IdTercero from Terceros where IdTercero='" & FufuSt & "'") = False Then
          Desbloquear
          limpiar
          TxtCampo(0).Text = FufuSt
          TxtCampo(13).SetFocus
          Editando = False
          TxtCampo(12).Text = 0
          CboFormaPago.ListIndex = 0
          FraNegociaciones.Enabled = False
          TxtCampo(14).Text = 1
          TxtCampo(15).Text = 1
          ChkContado.value = 1
          ChkDestino.value = 1
        Else
          MsgBox "Ya hay un tercero creado con esta identificacion, no se pueden crear dos datos con esta identificacion", vbCritical, "El tercero ya existe"
        End If
      End If
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            Bloquear
            AbrirRecorset rstUniversal, "Update terceros set TpDoc='" & TxtCampo(1).Text & "', RazonSocial='" & TxtCampo(2).Text & "', Nombre='" & TxtCampo(3).Text & "', Apellido1='" & TxtCampo(4).Text & "', Apellido2='" & TxtCampo(5).Text & "', Direccion='" & TxtCampo(6).Text & "', Telefono='" & TxtCampo(7).Text & "', IdCiudad=" & Val(TxtCampo(8).Text) & ", IdCliente=" & Val(TxtCampo(9).Text) & ", Email='" & TxtCampo(10).Text & "', Celular='" & TxtCampo(11).Text & "', Plazo = " & Val(TxtCampo(12).Text) & ", DigitoVerificacion= " & Val(TxtCampo(13).Text) & ", IdAsesor= " & Val(TxtCampo(14).Text) & ", IdCentroCostos= " & Val(TxtCampo(15).Text) & ", IdFormaPago = " & CboFormaPago.ListIndex + 1 & ", CondicionComercial = '" & TxtCampo(16).Text & "', ManejaCobroContado =" & ChkContado.value & ", ManejaCobroDestino=" & ChkDestino.value & ", ManejaCobroCorriente=" & ChkCorriente.value & " where IdTercero='" & TxtCampo(0).Text & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
            AccionTool 17
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          AbrirRecorset rstUniversal, "INSERT INTO Terceros (IdTercero, TpDoc, RazonSocial, Nombre, Apellido1, Apellido2, Direccion, Telefono, IdCiudad, IdCliente, Email, Celular, Plazo, DigitoVerificacion, IdAsesor, IdCentroCostos, IdFormaPago, CondicionComercial, ManejaCobroContado, ManejaCobroDestino, ManejaCobroCorriente) VALUES ('" & TxtCampo(0).Text & "', '" & TxtCampo(1).Text & "', '" & TxtCampo(2).Text & "', '" & TxtCampo(3).Text & "', '" & TxtCampo(4).Text & "', '" & TxtCampo(5).Text & "', '" & TxtCampo(6).Text & "', '" & TxtCampo(7).Text & "', " & Val(TxtCampo(8).Text) & ", " & Val(TxtCampo(9).Text) & ", '" & TxtCampo(10).Text & "', '" & TxtCampo(11).Text & "', " & Val(TxtCampo(12).Text) & ", " & Val(TxtCampo(13).Text) & ", " & Val(TxtCampo(14).Text) & ", " & Val(TxtCampo(15).Text) & ", " & CboFormaPago.ListIndex + 1 & ", '" & TxtCampo(16).Text & "', " & ChkContado.value & ", " & ChkDestino.value & ", " & ChkCorriente.value & " )", CnnPrincipal, adOpenDynamic, adLockOptimistic
          AccionTool 17
          Bloquear
        End If
      End If
    Case 5  'Editar
      Editando = True
      Desbloquear
    Case 6 'Eliminar
      EliminarTercero TxtCampo(0).Text
    Case 7 'Cancelar
      If Editando = True Then
        AccionTool 4
      Else
        Asignar rstTerceros
        Bloquear
      End If
    Case 9  'Buscar
      If Principal.ToolConsultas1.AbrirDevConsulta(7, CnnPrincipal) = True Then
        AbrirRecorset rstUniversal, strTerceros & " where IdTercero='" & Principal.ToolConsultas1.DatSt & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          Asignar rstUniversal
        Else
          MsgBox "No se encontraron terceros con este ID", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If
      
    Case 11 'Primero
      UPrimero rstTerceros
      Asignar rstTerceros
    Case 12 'Anterior
      UAnterior rstTerceros
      Asignar rstTerceros
    Case 13 'Siguiente
      USiguiente rstTerceros
      Asignar rstTerceros
    Case 14 'Ultimo
      UUltimo rstTerceros
      Asignar rstTerceros
    Case 16 'Cerrar
      CerrarRecorset rstTerceros
      FufuSt = TxtCampo(0)
      Unload Me
    Case 17 'Actualizar
      rstTerceros.Requery
    Case 18 'Imprimir
    Case 19
      If TxtCampo(9).Text <> "" Then TxtNmCliente.Text = DevResBus("SELECT Id, NmNegociacion From Negociaciones where Id=" & Val(TxtCampo(9)), "NmNegociacion", CnnPrincipal)
      If Val(TxtCampo(8).Text) <> 0 Then TxtNmCiudad.Text = DevResBus("SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & TxtCampo(8), "NmCiudad", CnnPrincipal)
      
  End Select
End Sub
Function Validacion() As Boolean
  Validacion = False
    If TxtCampo(1).Text <> "" Then
      If TxtCampo(13).Text <> "" Then
        If TxtCampo(2).Text <> "" Then
          If TxtCampo(6).Text <> "" Then
            If TxtCampo(8).Text <> "" Then
              If Val(TxtCampo(14).Text) <> 0 Then
                If Val(TxtCampo(15).Text) <> 0 Then
                  Validacion = True
                Else
                  MsgBox "El tercero debe tener un centro costos", vbCritical: TxtCampo(15).SetFocus
                End If
              Else
                MsgBox "El tercero debe tener un asesor", vbCritical: TxtCampo(14).SetFocus
              End If
            Else
              MsgBox "El tercero debe tener una ciudad", vbCritical: TxtCampo(8).SetFocus
            End If
          Else
            MsgBox "El tercero debe tener una direccion", vbCritical: TxtCampo(6).SetFocus
          End If
        Else
          MsgBox "El tercero debe tener un nombre extendido", vbCritical: TxtCampo(2).SetFocus
        End If
      Else
        MsgBox "El tercero debe tener un digito de verificacion", vbCritical: TxtCampo(13).SetFocus
      End If
    Else
      MsgBox "El tercero debe tener un tipo de identificacion", vbCritical: TxtCampo(1).SetFocus
    End If
End Function



Private Sub ToolTerceros_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub

Private Sub TxtCampo_GotFocus(Index As Integer)
  EnfocarT TxtCampo(Index)
  TxtCampo(Index).BackColor = &H80000001
  TxtCampo(Index).ForeColor = &HFFFFFF
End Sub

Private Sub TxtCampo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Select Case Index
      Case 8
        Principal.ToolConsultas1.AbrirDevConsulta 1, CnnPrincipal
        TxtCampo(8).Text = Principal.ToolConsultas1.DatLo
      Case 9
        Principal.ToolConsultas1.AbrirDevConsulta 2, CnnPrincipal
        TxtCampo(9).Text = Principal.ToolConsultas1.DatLo
      Case 14
        FrmBuscarAsesor.Show 1
        TxtCampo(14).Text = FufuLo
      Case 15
        FrmBuscarCentroCostos.Show 1
        TxtCampo(15).Text = FufuLo
        
    End Select
  End If
End Sub

Private Sub TxtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
  Select Case Index
    Case 1
      ValidarEntrada TxtCampo(1), KeyAscii, 4
    Case 7, 8, 9, 11, 14, 15
      ValidarEntrada TxtCampo(1), KeyAscii, 1
  End Select
End Sub
Private Sub TxtCampo_LostFocus(Index As Integer)
  TxtCampo(Index).BackColor = &H80000005
  TxtCampo(Index).ForeColor = &H80000012
End Sub

Private Sub TxtCampo_Validate(Index As Integer, Cancel As Boolean)
  Select Case Index
    Case 8
      If Val(TxtCampo(8).Text) <> 0 Then
        AbrirRecorset rstUniversal, "SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & TxtCampo(8), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          TxtNmCiudad.Text = rstUniversal!NmCiudad & ""
        Else
          TxtNmCiudad.Text = "": TxtCampo(8).Text = ""
        End If
        CerrarRecorset rstUniversal
      End If
      
    Case 9
      AbrirRecorset rstUniversal, "Select Id, NmNegociacion from Negociaciones where Id=" & Val(TxtCampo(9)), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        TxtNmCliente.Text = rstUniversal.Fields("NmNegociacion") & ""
      Else
        TxtNmCliente.Text = "": TxtCampo(9).Text = ""
      End If
      CerrarRecorset rstUniversal
      
    Case 14
      If Val(TxtCampo(14).Text) <> 0 Then
        AbrirRecorset rstUniversal, "SELECT IdAsesor, NmAsesor From Asesores where IdAsesor=" & Val(TxtCampo(14).Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          TxtNmAsesor.Text = rstUniversal!NmAsesor & ""
        Else
          TxtNmAsesor.Text = "": TxtCampo(14).Text = ""
        End If
        CerrarRecorset rstUniversal
      Else
        TxtCampo(14).Text = ""
        TxtNmAsesor.Text = ""
      End If
      
    Case 15
      If Val(TxtCampo(15).Text) <> 0 Then
        AbrirRecorset rstUniversal, "SELECT IdCentroCostos, NmCentroCostos From centros_costos where IdCentroCostos=" & Val(TxtCampo(15).Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          TxtNmCentroCostos.Text = rstUniversal!NmCentroCostos & ""
        Else
          TxtNmCentroCostos.Text = "": TxtCampo(15).Text = ""
        End If
        CerrarRecorset rstUniversal
      Else
        TxtCampo(15).Text = ""
        TxtNmCentroCostos.Text = ""
      End If
  End Select
End Sub
Private Sub JuntarNombre()
  TxtCampo(2).Text = ""
  If TxtCampo(4).Text <> "" Then
    TxtCampo(2) = TxtCampo(2).Text & TxtCampo(4) & " "
  End If
  If TxtCampo(5).Text <> "" Then
    TxtCampo(2) = TxtCampo(2).Text & TxtCampo(5) & " "
  End If
  If TxtCampo(3).Text <> "" Then
    TxtCampo(2) = TxtCampo(2).Text & TxtCampo(3)
  End If
End Sub

Private Sub EliminarTercero(IdTercero As String)
Dim rstTercero As New ADODB.Recordset
rstTercero.CursorLocation = adUseClient
On Error GoTo ElError
  If MsgBox("¿Esta seguro de eliminar este nit?", vbQuestion + vbYesNo, "Eliminar Nit") = vbYes Then
    rstTercero.Open "Delete from Terceros where IdTercero='" & TxtCampo(0).Text & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
    MsgBox "Tercero eliminaro correctamente", vbInformation
    AccionTool 17
    AccionTool 11
  End If
ElError:
  If Err.Number <> 0 Then
    MsgBox "No se puede eliminar el tercero porque hay guias o facturas con este nit", vbCritical
  End If
End Sub

Private Sub VerNegociaciones()
  Dim strSql As String
  strSql = "SELECT negociaciones_terceros.*, negociaciones.NmNegociacion " & _
           "FROM negociaciones_terceros " & _
           "LEFT JOIN negociaciones ON negociaciones_terceros.IdNegociacion = negociaciones.Id " & _
           "WHERE IdTercero = '" & TxtCampo(0).Text & "'"
           
  LstNegociaciones.ListItems.Clear
  AbrirRecorset rstUniversal, strSql, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstUniversal.RecordCount > 0 Then
    Do While rstUniversal.EOF = False
      Set Item = LstNegociaciones.ListItems.Add(, , rstUniversal!IdNegociacion)
      Item.SubItems(1) = rstUniversal!NmNegociacion & ""
      rstUniversal.MoveNext
    Loop
  End If
  CerrarRecorset rstUniversal
End Sub
