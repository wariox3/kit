VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLlenarFacturas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Llenar factura..."
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12390
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   12390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Orden"
      Height          =   735
      Left            =   11160
      TabIndex        =   24
      Top             =   4080
      Width           =   1095
      Begin VB.OptionButton Option2 
         Caption         =   "Guia"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   450
         Width           =   855
      End
      Begin VB.OptionButton OptOrdenFecha 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   220
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdCambiarNroDoc 
      Caption         =   "Cambiar documento"
      Height          =   255
      Left            =   7320
      TabIndex        =   23
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CheckBox ChkNegociacion 
      Caption         =   "Negociacion"
      Height          =   255
      Left            =   9600
      TabIndex        =   21
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton CmdCambiarNegociacion 
      Caption         =   "Cambiar"
      Height          =   255
      Left            =   10200
      TabIndex        =   20
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox TxtIdNegociacion 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9600
      TabIndex        =   19
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton CmdAgregarDocumento 
      Caption         =   "Agregar por doc >>"
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9600
      TabIndex        =   12
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   "Guias facturadas"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9600
      TabIndex        =   10
      Top             =   4800
      Width           =   2655
      Begin VB.TextBox TxtNroGuias 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Text            =   "0"
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Totales"
      Enabled         =   0   'False
      Height          =   975
      Left            =   9600
      TabIndex        =   5
      Top             =   5520
      Width           =   2655
      Begin VB.TextBox TxtTotalFlete 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Text            =   "0"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TxtTotalManejo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Text            =   "0"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Flete:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   390
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Manejo:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   570
      End
   End
   Begin VB.CommandButton CmdQuitarMarcadas 
      Caption         =   "<< Quitar marcadas"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   6720
      Width           =   2655
   End
   Begin MSComctlLib.ListView LstTem 
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4260
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Guia"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Vr Flete"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Vr Manejo"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton CmdAgregarUaU 
      Caption         =   "Agregar por guia >>"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton CmdAgregarSel 
      Caption         =   "Agregar >>"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton CmdVerPendientes 
      Caption         =   "Ver pendientes"
      Height          =   255
      Left            =   9600
      TabIndex        =   1
      Top             =   3720
      Width           =   2655
   End
   Begin MSDataGridLib.DataGrid GrillaPendientes 
      Height          =   3255
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   5741
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
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "Guia"
         Caption         =   "Guia"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "DocCliente"
         Caption         =   "Documento"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "FhEntradaBodega"
         Caption         =   "Fh Entrada"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "COIng"
         Caption         =   "CO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "NmCiudad"
         Caption         =   "Destino"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Unidades"
         Caption         =   "Und"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "KilosReales"
         Caption         =   "K Real"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "KilosVolumen"
         Caption         =   "K Vol"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "KilosFacturados"
         Caption         =   "K Fac"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "VrDeclarado"
         Caption         =   "Declara"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "IdTpCtaFlete"
         Caption         =   "CF"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "VrFlete"
         Caption         =   "Flete"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "IdTpCtaManejo"
         Caption         =   "CM"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "VrManejo"
         Caption         =   "Manejo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column10 
            Locked          =   -1  'True
            ColumnWidth     =   329.953
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column12 
            Locked          =   -1  'True
            ColumnWidth     =   329.953
         EndProperty
         BeginProperty Column13 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   12360
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label LblNroRegistros 
      Caption         =   "Ver Pendientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   5040
      TabIndex        =   22
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label LblIdCliente 
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
      TabIndex        =   16
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label LblNmCliente 
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
      Left            =   1320
      TabIndex        =   15
      Top             =   120
      Width           =   9015
   End
   Begin VB.Label LblNroFactura 
      Alignment       =   1  'Right Justify
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
      Left            =   11520
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Factura:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10680
      TabIndex        =   13
      Top             =   120
      Width           =   720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   12360
      Y1              =   6600
      Y2              =   6600
   End
End
Attribute VB_Name = "FrmLlenarFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstTem As New ADODB.Recordset

Private Sub CmdAgregarDocumento_Click()
  If rstTem.State = adStateOpen Then
    Do While Principal.ToolConsultas1.AbrirDevDatos("Documento del cliente", "Digite el numero del documento a buscar", 2, 0) = True
      If BuscaRegistro("DocCliente='" & Principal.ToolConsultas1.DatSt & "'", rstTem) = True Then
        CmdAgregarSel_Click
      Else
        MsgBox "No se encontro la guia con este numero de documento " & Principal.ToolConsultas1.DatSt, vbCritical
      End If
    Loop
  End If
End Sub

Private Sub CmdAgregarSel_Click()
If rstTem.State = adStateOpen Then
  If rstTem.EOF = False Then
    AgregarGuia rstTem!Guia, rstTem!VrFlete, rstTem!VrManejo
    'USiguiente rstTem
  End If
Else
  If MsgBox("Debe ver los pendientes por facturar de este cliente para seleccionar el registro a facturar" & Chr(13) & "¿Desea ver los pendientes por facturar de este cliente?", vbQuestion + vbYesNo) = vbYes Then CmdVerPendientes_Click
End If
End Sub

Private Sub CmdAgregarUaU_Click()
  If rstTem.State = adStateOpen Then
    Do While Principal.ToolConsultas1.AbrirDevDatos("Numero de Guia", "Digite el numero de la guia que desea buscar", 3, 0) = True
      If BuscaRegistro("Guia=" & Principal.ToolConsultas1.DatLo, rstTem) = True Then
        CmdAgregarSel_Click
      Else
        MsgBox "No se encontro la guia con numero " & Principal.ToolConsultas1.DatLo, vbCritical
      End If
    Loop
  End If
End Sub

Private Sub CmdCambiarNegociacion_Click()
  FufuSt = LblIdCliente.Caption
  FrmBuscarNegociaciones.Show 1
  If FufuLo <> 0 Then
    TxtIdNegociacion.Text = FufuLo
  End If
End Sub

Private Sub CmdCambiarNroDoc_Click()
Dim NuevoDoc As String
If rstTem.State <> adStateOpen Then MsgBox "No ha actualizado", vbCritical: Exit Sub
  If rstTem.EOF = False Then
    If MsgBox("Esta seguro que desea cambiarle el documento a la Guia " & rstTem.Fields("Guia") & " con Documento " & rstTem.Fields("DocCliente"), vbInformation + vbYesNo) = vbYes Then
      NuevoDoc = InputBox("Digite el nuevo documento del cliente")
      If Len(NuevoDoc) > 0 And Len(NuevoDoc) < 20 Then
        AbrirRecorset rstUniversal, "Update guias set DocCliente='" & NuevoDoc & "' where guia=" & rstTem.Fields("guia"), CnnPrincipal, adOpenDynamic, adLockOptimistic
      Else
        MsgBox "El nuevo documento debe contener al menos 1 caracter y menos de 20", vbCritical
      End If
      
    End If
  End If
End Sub

Private Sub CmdGuardar_Click()
  Set rstTem = Nothing
  AbrirRecorset rstUniversal, "Update Guias set IdFactura=0, Facturada=0 where IdFactura=" & Val(LblNroFactura.Caption), CnnPrincipal, adOpenDynamic, adLockOptimistic

  IniProg (LstTem.ListItems.Count)
  For II = 1 To LstTem.ListItems.Count
    AbrirRecorset rstUniversal, "Update Guias set IdFactura=" & Val(LblNroFactura.Caption) & ", Facturada=1 where Guia=" & LstTem.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
    Prog (II)
  Next II
  FinProg
  FrmFacturas.TxtCampos(17) = TxtNroGuias
  FrmFacturas.TxtCampos(6) = TxtTotalFlete
  FrmFacturas.TxtCampos(7) = TxtTotalManejo
  Unload Me
End Sub
Private Sub CmdQuitarMarcadas_Click()
II = 1
Do While II <= LstTem.ListItems.Count
  If LstTem.ListItems(II).Checked = True Then
    TxtTotalFlete = TxtTotalFlete - LstTem.ListItems(II).SubItems(1)
    TxtTotalManejo = TxtTotalManejo - LstTem.ListItems(II).SubItems(2)
    TxtNroGuias = Val(TxtNroGuias) - 1
    LstTem.ListItems.Remove (II)
  Else
    II = II + 1
  End If
Loop
End Sub

Private Sub CmdVerPendientes_Click()
  If ChkNegociacion.Value = 1 And Val(TxtIdNegociacion.Text) Then
    AbrirRecorset rstTem, "SELECT * from sql_if_pend_fac Where Cuenta='" & LblIdCliente.Caption & "' and IdCliente=" & Val(TxtIdNegociacion) & DevOrdenPendientes(), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Else
    AbrirRecorset rstTem, "SELECT * from sql_if_pend_fac Where Cuenta='" & LblIdCliente.Caption & "'" & DevOrdenPendientes(), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  End If
  Set GrillaPendientes.DataSource = rstTem
  LblNroRegistros.Caption = rstTem.RecordCount & " Registros"
End Sub
Private Sub AgregarGuia(NroGuia As String, VrFlete As Currency, VrManejo As Currency)
    Set Item = LstTem.FindItem(NroGuia)
    If Item Is Nothing Then
      Set Item = LstTem.ListItems.Add(, , NroGuia)
        Item.SubItems(1) = VrFlete
        Item.SubItems(2) = VrManejo
        TxtTotalFlete = Format(TxtTotalFlete + VrFlete, "#,##0.00;(#,##0.00)")
        TxtTotalManejo = Format(TxtTotalManejo + VrManejo, "#,##0.00;(#,##0.00)")
        TxtNroGuias = Val(TxtNroGuias) + 1
    Else
      MsgBox "La guia [" & NroGuia & "] ya se le agrego al temporal para facturar", vbCritical, "La guia ya fue agregada"
    End If
End Sub



Private Sub Form_Load()
  LblNmCliente = FrmFacturas.TxtNmCliente.Text
  LblIdCliente = FrmFacturas.TxtCampos(3)
  LblNroFactura.Caption = FrmFacturas.TxtCampos(0)
  rstTem.CursorLocation = adUseClient
  AbrirRecorset rstUniversal, "SELECT Guia, VrFlete, VrManejo, IdTpCtaFlete, IdTpCtaManejo, IdFactura " & _
                              "FROM Guias " & _
                              "WHERE IdFactura = " & Val(LblNroFactura.Caption), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    AgregarGuia rstUniversal!Guia, rstUniversal!VrFlete, rstUniversal!VrManejo
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstTem = Nothing
End Sub

Private Sub GrillaPendientes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then CmdAgregarSel_Click
End Sub

Private Function DevOrdenPendientes() As String
  If OptOrdenFecha.Value = True Then
    DevOrdenPendientes = " ORDER BY FhEntradaBodega"
  Else
    DevOrdenPendientes = " ORDER BY Guia"
  End If
End Function
