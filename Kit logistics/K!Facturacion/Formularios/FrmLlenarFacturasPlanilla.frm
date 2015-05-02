VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLlenarFacturasPlanilla 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Llenar factura por planilla"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   6120
      TabIndex        =   36
      Top             =   6850
      Width           =   1695
      Begin VB.OptionButton OptImportarPorGuia 
         Caption         =   "Guia"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton OptImportarPorDocumento 
         Caption         =   "Documento"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdImportar 
      Caption         =   "Importar"
      Height          =   255
      Left            =   6120
      TabIndex        =   35
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox TxtRuta 
      Height          =   285
      Left            =   2280
      TabIndex        =   34
      Text            =   "C:\pruebaimportacion.csv"
      Top             =   6600
      Width           =   3135
   End
   Begin VB.CommandButton CmdSeleccionar 
      Caption         =   "..."
      Height          =   255
      Left            =   5520
      TabIndex        =   33
      Top             =   6600
      Width           =   495
   End
   Begin VB.CheckBox ChkRelCliente 
      Caption         =   "Rel Cliente"
      Height          =   195
      Left            =   8640
      TabIndex        =   32
      Top             =   3890
      Width           =   1335
   End
   Begin VB.CommandButton CmdCambiarNroDoc 
      Caption         =   "Cambiar documento"
      Height          =   255
      Left            =   7920
      TabIndex        =   30
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton CmdAgregarDocumento 
      Caption         =   "Agregar por doc >>"
      Height          =   255
      Left            =   3720
      TabIndex        =   29
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox TxtIdNegociacion 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7080
      TabIndex        =   23
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton CmdCambiarNegociacion 
      Caption         =   "Cambiar"
      Height          =   255
      Left            =   7680
      TabIndex        =   22
      Top             =   3720
      Width           =   855
   End
   Begin VB.CheckBox ChkNegociacion 
      Caption         =   "Negociacion"
      Height          =   255
      Left            =   8640
      TabIndex        =   21
      Top             =   3650
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Totales planilla"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   10200
      TabIndex        =   14
      Top             =   5160
      Width           =   2175
      Begin VB.TextBox TxtFletePlanilla 
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
         Left            =   720
         TabIndex        =   17
         Text            =   "0"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtManejoPlanilla 
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
         Left            =   720
         TabIndex        =   16
         Text            =   "0"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxtGuiasPlanilla 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   720
         TabIndex        =   15
         Text            =   "0"
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Flete:"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   390
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Manejo:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   570
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Guias:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   450
      End
   End
   Begin VB.CommandButton CmdVerPendientes 
      Caption         =   "Ver pendientes"
      Height          =   255
      Left            =   10080
      TabIndex        =   10
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton CmdAgregarSel 
      Caption         =   "Agregar >>"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton CmdAgregarUaU 
      Caption         =   "Agregar por guia >>"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton CmdQuitarMarcadas 
      Caption         =   "<< Quitar marcadas"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Totales"
      Enabled         =   0   'False
      Height          =   975
      Left            =   10200
      TabIndex        =   1
      Top             =   4080
      Width           =   2175
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
         Left            =   720
         TabIndex        =   3
         Text            =   "0"
         Top             =   600
         Width           =   1335
      End
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
         Left            =   720
         TabIndex        =   2
         Text            =   "0"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Manejo:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   570
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Flete:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   390
      End
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "&Aceptar"
      Height          =   255
      Left            =   10200
      TabIndex        =   0
      Top             =   6600
      Width           =   2175
   End
   Begin MSComctlLib.ListView LstTem 
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   9975
      _ExtentX        =   17595
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
   Begin MSDataGridLib.DataGrid GrillaPendientes 
      Height          =   3255
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   12255
      _ExtentX        =   21616
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
      ColumnCount     =   15
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
      BeginProperty Column14 
         DataField       =   "RelCliente"
         Caption         =   "Relacion"
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
            Locked          =   -1  'True
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
         BeginProperty Column14 
         EndProperty
      EndProperty
   End
   Begin VB.Label LblRelCliente 
      Height          =   255
      Left            =   7920
      TabIndex        =   31
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label LblNroPlanilla 
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
      Left            =   9840
      TabIndex        =   28
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Planilla:"
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
      Left            =   9120
      TabIndex        =   27
      Top             =   120
      Width           =   690
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
      TabIndex        =   26
      Top             =   120
      Width           =   720
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
      TabIndex        =   25
      Top             =   120
      Width           =   735
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
      Left            =   5520
      TabIndex        =   24
      Top             =   3720
      Width           =   1455
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
      TabIndex        =   13
      Top             =   120
      Width           =   5895
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
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "FrmLlenarFacturasPlanilla"
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
    AgregarGuia rstTem!Guia, rstTem!VrFlete, rstTem!VrManejo, True, Val(rstTem!IdPlanillaFactura & "")
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
  AbrirRecorset rstUniversal, "Update Guias set IdFactura=0, Facturada=0, IdPlanillaFactura=0 where IdFactura=" & Val(LblNroFactura.Caption) & " and IdPlanillaFactura=" & Val(LblNroPlanilla.Caption), CnnPrincipal, adOpenDynamic, adLockOptimistic
  IniProg (LstTem.ListItems.Count)
  For II = 1 To LstTem.ListItems.Count
    AbrirRecorset rstUniversal, "Update Guias set IdFactura=" & Val(LblNroFactura.Caption) & ", Facturada=1, IdPlanillaFactura=" & Val(LblNroPlanilla.Caption) & " where Guia=" & LstTem.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
    Prog (II)
  Next II
  FinProg
  FrmFacturas.TxtCampos(6) = TxtTotalFlete
  FrmFacturas.TxtCampos(7) = TxtTotalManejo
  AbrirRecorset rstUniversal, "Update FacturasPlanillas set VrFletePlanilla=" & Val(Format(TxtFletePlanilla.Text, "0;(0)")) & ", VrManejoPlanilla=" & Val(Format(TxtManejoPlanilla.Text, "0;(0)")) & ", NroGuiasPlanilla=" & Val(Format(TxtGuiasPlanilla.Text, "0;(0)")) & " where IdPlanilla=" & Val(LblNroPlanilla), CnnPrincipal, adOpenDynamic, adLockOptimistic
  Unload Me
End Sub


Private Sub CmdImportar_Click()
  Dim Numero As String
  Open TxtRuta.Text For Input As #1
  Do While Not EOF(1)
    Input #1, Numero
    If OptImportarPorDocumento.Value = True Then
      If BuscaRegistro("DocCliente=" & Numero, rstTem) = True Then
        CmdAgregarSel_Click
      End If
    Else
      If BuscaRegistro("Guia=" & Numero, rstTem) = True Then
        CmdAgregarSel_Click
      End If
    End If
  Loop
  Close #1
End Sub

Private Sub CmdQuitarMarcadas_Click()
II = 1
Do While II <= LstTem.ListItems.Count
  If LstTem.ListItems(II).Checked = True Then
      TxtTotalFlete = TxtTotalFlete - LstTem.ListItems(II).SubItems(1)
      TxtFletePlanilla = TxtFletePlanilla - LstTem.ListItems(II).SubItems(1)
      TxtTotalManejo = TxtTotalManejo - LstTem.ListItems(II).SubItems(2)
      TxtManejoPlanilla = TxtManejoPlanilla - LstTem.ListItems(II).SubItems(2)
      TxtGuiasPlanilla = Val(TxtGuiasPlanilla) - 1
      LstTem.ListItems.Remove (II)
  Else
    II = II + 1
  End If
Loop
End Sub

Private Sub CmdSeleccionar_Click()
  Principal.CDExa.ShowOpen
  If Principal.CDExa.FileName <> "" Then
   TxtRuta.Text = Principal.CDExa.FileName
  End If
End Sub

Private Sub CmdVerPendientes_Click()
  Dim sql As String
  If ChkRelCliente.Value = 1 Then
    sql = sql & " and RelCliente='" & LblRelCliente.Caption & "'"
  End If

  If ChkNegociacion.Value = 1 And Val(TxtIdNegociacion.Text) Then
    sql = sql & " and IdCliente=" & Val(TxtIdNegociacion)
  End If
  
  
  AbrirRecorset rstTem, "SELECT sql_if_pend_fac.* from sql_if_pend_fac Where Cuenta='" & LblIdCliente.Caption & "' " & sql & " ORDER BY FhEntradaBodega", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  
  Set GrillaPendientes.DataSource = rstTem
  LblNroRegistros.Caption = rstTem.RecordCount & " Registros"
End Sub
Private Sub AgregarGuia(NroGuia As String, VrFlete As Currency, VrManejo As Currency, Totales As Boolean, IdPlanilla As Long)
  Set Item = LstTem.FindItem(NroGuia)
  If Item Is Nothing Then
    Set Item = LstTem.ListItems.Add(, , NroGuia)
      Item.SubItems(1) = VrFlete
      Item.SubItems(2) = VrManejo
      If Totales = True Then
        TxtTotalFlete = Format(TxtTotalFlete + VrFlete, "#,##0.00;(#,##0.00)")
        TxtTotalManejo = Format(TxtTotalManejo + VrManejo, "#,##0.00;(#,##0.00)")
      End If
      
      TxtFletePlanilla = Format(TxtFletePlanilla + VrFlete, "#,##0.00;(#,##0.00)")
      TxtManejoPlanilla = Format(TxtManejoPlanilla + VrManejo, "#,##0.00;(#,##0.00)")
      TxtGuiasPlanilla = Val(TxtGuiasPlanilla) + 1
      
  Else
    MsgBox "La guia [" & NroGuia & "] ya se le agrego al temporal para facturar", vbCritical, "La guia ya fue agregada"
  End If
End Sub





Private Sub Form_Load()
  LblNmCliente = FrmFacturas.TxtNmCliente.Text
  LblIdCliente = FrmFacturas.TxtCampos(3)
  LblNroFactura.Caption = FrmFacturas.TxtCampos(0).Text
  LblNroPlanilla.Caption = FufuLo
  AbrirRecorset rstUniversal, "Select IdPlanilla, RelCliente from facturasplanillas where IdPlanilla=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.RecordCount > 0 Then
    LblRelCliente.Caption = rstUniversal.Fields("RelCliente")
  End If
  CerrarRecorset rstUniversal
  
  rstTem.CursorLocation = adUseClient
  AbrirRecorset rstUniversal, "SELECT Guia, VrFlete, VrManejo, IdTpCtaFlete, IdTpCtaManejo, IdFactura, IdPlanillaFactura From Guias where IdFactura=" & Val(LblNroFactura.Caption) & " and IdPlanillaFactura=" & Val(LblNroPlanilla.Caption), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    AgregarGuia rstUniversal!Guia, rstUniversal!VrFlete, rstUniversal!VrManejo, False, rstUniversal!IdPlanillaFactura
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
  
  AbrirRecorset rstUniversal, "Select*from facturas where IdFactura=" & Val(LblNroFactura.Caption), CnnPrincipal, adOpenDynamic, adLockOptimistic
    TxtTotalFlete = Format(rstUniversal.Fields("TFlete"), "#,##0.00;(#,##0.00)")
    TxtTotalManejo = Format(rstUniversal.Fields("TManejo"), "#,##0.00;(#,##0.00)")
  CerrarRecorset rstUniversal
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstTem = Nothing
End Sub

Private Sub GrillaPendientes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then CmdAgregarSel_Click
End Sub


