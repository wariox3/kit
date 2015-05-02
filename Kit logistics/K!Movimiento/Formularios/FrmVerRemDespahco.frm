VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmVerGuiasDespacho 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ver guias del despacho..."
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14010
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   14010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCambiar 
      Caption         =   "Cambiar"
      Height          =   255
      Left            =   2640
      TabIndex        =   28
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox TxtDespacho 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   27
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame FraCriterio 
      Caption         =   "Criterio"
      Height          =   1335
      Left            =   11280
      TabIndex        =   23
      Top             =   5040
      Width           =   2535
      Begin VB.OptionButton OptCriterio 
         Caption         =   "Sin entregar"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton OptCriterio 
         Caption         =   "Descargadas"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton OptCriterio 
         Caption         =   "Sin descargar"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton OptCriterio 
         Caption         =   "Todos"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   11055
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   1
         Left            =   8880
         TabIndex        =   4
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   2
         Left            =   5760
         TabIndex        =   5
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   3
         Left            =   5760
         TabIndex        =   6
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   4
         Left            =   960
         TabIndex        =   7
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   6
         Left            =   960
         TabIndex        =   9
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   7
         Left            =   5760
         TabIndex        =   10
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   8
         Left            =   3240
         TabIndex        =   18
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTotales 
         Height          =   255
         Index           =   9
         Left            =   3240
         TabIndex        =   20
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "K Fac:"
         Height          =   195
         Index           =   9
         Left            =   2760
         TabIndex        =   21
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "K Vol:"
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   19
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Manejo:"
         Height          =   195
         Index           =   8
         Left            =   5160
         TabIndex        =   17
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Flete:"
         Height          =   195
         Index           =   4
         Left            =   5280
         TabIndex        =   16
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cobro Destino:"
         Height          =   195
         Index           =   6
         Left            =   7800
         TabIndex        =   15
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remesas:"
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   14
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unidades:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "K Reales:"
         Height          =   195
         Index           =   2
         Left            =   2520
         TabIndex        =   12
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Declarado:"
         Height          =   195
         Index           =   0
         Left            =   4920
         TabIndex        =   11
         Top             =   960
         Width           =   780
      End
   End
   Begin MSDataGridLib.DataGrid GrillaGuias 
      Height          =   4335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   7646
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
      ColumnCount     =   18
      BeginProperty Column00 
         DataField       =   "Guia"
         Caption         =   "Guia"
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
         DataField       =   "COing"
         Caption         =   "CO"
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
         DataField       =   "CR"
         Caption         =   "CR"
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
      BeginProperty Column03 
         DataField       =   "FhEntradaBodega"
         Caption         =   "FH Ent"
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
      BeginProperty Column04 
         DataField       =   "IdCliente"
         Caption         =   "ID Cuenta"
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
      BeginProperty Column05 
         DataField       =   "Remitente"
         Caption         =   "Remitente"
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
      BeginProperty Column06 
         DataField       =   "DocCliente"
         Caption         =   "Doc"
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
      BeginProperty Column07 
         DataField       =   "NmDestinatario"
         Caption         =   "Destinatario"
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
      BeginProperty Column08 
         DataField       =   "NmCiudad"
         Caption         =   "Destino"
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
      BeginProperty Column09 
         DataField       =   "VrDeclarado"
         Caption         =   "Declarado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "VrFlete"
         Caption         =   "Flete"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "VrManejo"
         Caption         =   "Manejo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Unidades"
         Caption         =   "Unidades"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "KilosReales"
         Caption         =   "KReales"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "KilosFacturados"
         Caption         =   "K Fac"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "KilosVolumen"
         Caption         =   "K Vol"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "TpServicio"
         Caption         =   "Tp S"
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
      BeginProperty Column17 
         DataField       =   "CPorte"
         Caption         =   "CP"
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
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   1860.095
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column08 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column13 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column15 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column16 
            Locked          =   -1  'True
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column17 
            Locked          =   -1  'True
            ColumnWidth     =   360
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdVer 
      Caption         =   "Ver / Actualizar"
      Height          =   255
      Left            =   9840
      TabIndex        =   1
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Aceptar / Salir"
      Height          =   255
      Left            =   11880
      TabIndex        =   0
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Despacho:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   22
      Top             =   4560
      Width           =   1140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   120
      X2              =   13800
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      Index           =   0
      X1              =   13800
      X2              =   120
      Y1              =   4920
      Y2              =   4920
   End
End
Attribute VB_Name = "FrmVerGuiasDespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstTemp As New ADODB.Recordset
Sub VerTotales()
    AbrirRecorset rstUniversal, "SELECT SUM(VrFlete+VrManejo-Abonos) as TCE FROM Guias WHERE IdDespacho=" & Val(TxtDespacho.Text) & "  and TipoCobro = 2", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      TxtTotales(1) = rstUniversal!TCE & ""
    AbrirRecorset rstUniversal, "SELECT SUM(Unidades) as TUnidades, Sum(KilosReales) as TKReales, sum(KilosVolumen) as TKVol, sum(KilosFacturados) as TKFac, Sum(VrFlete) as TFlete, Sum(VrManejo) as TManejo, Sum(VrDeclarado) as TDeclarado, Count(*) as NroReg FROM Guias WHERE IdDespacho=" & Val(TxtDespacho.Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly

    TxtTotales(2) = rstUniversal!TManejo & ""
    TxtTotales(3) = rstUniversal!TFlete & ""
    TxtTotales(4) = rstUniversal!TUnidades & ""
    TxtTotales(5) = rstUniversal!TKReales & ""
    TxtTotales(6) = rstUniversal!NroReg & ""
    TxtTotales(7) = rstUniversal!TDeclarado & ""
    TxtTotales(8) = rstUniversal!TKVol & ""
    TxtTotales(9) = rstUniversal!TKFac & ""
End Sub

Private Sub CmdCambiar_Click()
  If Principal.ToolConsultas1.AbrirDevDatos("Digite el numero del despacho", "Digite el numero del despacho", 3, 0) = True Then
    AbrirRecorset rstUniversal, "SELECT OrdDespacho FROM Despachos WHERE OrdDespacho = " & Principal.ToolConsultas1.DatLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
    If rstUniversal.RecordCount > 0 Then
      TxtDespacho.Text = Principal.ToolConsultas1.DatLo
      CmdVer_Click
    End If
  End If
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub
Private Sub CmdVer_Click()
  Dim Sql As String
  Sql = "SELECT Guias.Guia, Guias.CR, Guias.Remitente, Guias.IdCliente, Guias.DocCliente, Guias.NmDestinatario, Ciudades.NmCiudad, Guias.FhEntradaBodega, Guias.VrDeclarado, Guias.VrFlete, Guias.VrManejo, Guias.Unidades, Guias.KilosReales, Guias.KilosFacturados, Guias.KilosVolumen, Guias.Estado, Guias.IdDespacho, Guias.IdFactura, Guias.Observaciones, Guias.COIng, Guias.TpServicio, Guias.IdTpCtaFlete, Guias.IdTpCtaManejo, Guias.CPorte FROM Guias INNER JOIN Ciudades ON Guias.IdCiuDestino = Ciudades.IdCiudad where IdDespacho=" & Val(TxtDespacho.Text)
  If Val(TxtDespacho.Text) <> 0 Then
    If OptCriterio(0).value = True Then
      AbrirRecorset rstTemp, Sql, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    ElseIf OptCriterio(1).value = True Then
      AbrirRecorset rstTemp, Sql & " and Descargada=0", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    ElseIf OptCriterio(2).value = True Then
      AbrirRecorset rstTemp, Sql & " and Descargada=1", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    ElseIf OptCriterio(3).value = True Then
      AbrirRecorset rstTemp, Sql & " and Entregada=0", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    End If
    Set GrillaGuias.DataSource = rstTemp
    VerTotales
  End If
End Sub



Private Sub Form_Load()
  rstTemp.CursorLocation = adUseClient
  TxtDespacho.Text = FufuLo
  VerTotales
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstTemp = Nothing
End Sub
