VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCuentasCobrar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas por cobrar"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   13590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdVerdetalles 
      Caption         =   "Detalles"
      Height          =   375
      Left            =   7800
      TabIndex        =   9
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox TxtNumero 
      Height          =   315
      Left            =   3120
      TabIndex        =   6
      Top             =   7200
      Width           =   1455
   End
   Begin VB.ComboBox CboTipo 
      Height          =   315
      ItemData        =   "FrmCuentasCobrar.frx":0000
      Left            =   720
      List            =   "FrmCuentasCobrar.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton CmdExportar 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   9240
      TabIndex        =   3
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton CmdFiltrar 
      Caption         =   "Filtrar"
      Height          =   375
      Left            =   10680
      TabIndex        =   2
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   12120
      TabIndex        =   0
      Top             =   7200
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid Grilla 
      Height          =   6855
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   12091
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "IdCxC"
         Caption         =   "ID"
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
         DataField       =   "NmTipoFactura"
         Caption         =   "Tipo"
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
         DataField       =   "NroDocumento"
         Caption         =   "Numero"
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
      BeginProperty Column03 
         DataField       =   "FechaDoc"
         Caption         =   "Fecha"
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
         DataField       =   "RazonSocial"
         Caption         =   "Cliente"
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
         DataField       =   "NmAsesor"
         Caption         =   "Asesor"
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
      BeginProperty Column06 
         DataField       =   "Condicion"
         Caption         =   "Plazo"
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
      BeginProperty Column07 
         DataField       =   "DiasVencida"
         Caption         =   "Dias.V"
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
      BeginProperty Column08 
         DataField       =   "Saldo"
         Caption         =   "Saldo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
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
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2865.26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1934.929
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1200.189
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Numero:"
      Height          =   195
      Left            =   2520
      TabIndex        =   7
      Top             =   7200
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   7200
      Width           =   360
   End
End
Attribute VB_Name = "FrmCuentasCobrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCuentasCobrar As New ADODB.Recordset

Private Sub CmdBuscar_Click()
    If Val(TxtNumero.Text) <> 0 Then
      Dim strSql As String
      strSql = "Select cuentas_cobrar.*, NmTipoFactura, NmAsesor, NmPuntoOperaciones, RazonSocial " & _
                "from cuentas_cobrar " & _
                "left join terceros on cuentas_cobrar.IdTercero = terceros.IDTercero " & _
                "left join facturas_tipos on cuentas_cobrar.TipoFactura = facturas_tipos.IdTipoFactura " & _
                "left join asesores on cuentas_cobrar.IdAsesor = asesores.IdAsesor " & _
                "left join centrosoperaciones on cuentas_cobrar.IdPO = centrosoperaciones.IDPO " & _
                "where NroDocumento = " & TxtNumero.Text & " "
      If CboTipo.ListIndex <> 0 Then
        strSql = strSql & "and TipoFactura = " & CboTipo.ListIndex & ""
      End If
      AbrirRecorset rstCuentasCobrar, strSql, CnnPrincipal, adOpenStatic, adLockReadOnly
      Set Grilla.DataSource = rstCuentasCobrar
    End If
End Sub

Private Sub CmdExportar_Click()
If rstCuentasCobrar.State = adStateOpen Then
  If rstCuentasCobrar.EOF = False Then
    ExportarExcel rstCuentasCobrar
  End If
End If
End Sub

Private Sub CmdFiltrar_Click()
  FrmInformeCarteraEdades.Show 1
  If varParametrosCartera.Generar = True Then
    AbrirRecorset rstCuentasCobrar, varParametrosCartera.sql, CnnPrincipal, adOpenStatic, adLockReadOnly
    Set Grilla.DataSource = rstCuentasCobrar
  End If
End Sub

Private Sub CmdSalir_Click()
  Set rstCuentasCobrar = Nothing
  Unload Me
End Sub

Private Sub CmdVerdetalles_Click()
  If rstCuentasCobrar.State = adStateOpen Then
    If rstCuentasCobrar.EOF = False Then
      FufuLo = rstCuentasCobrar.Fields("IdCxC")
      FrmVerCuentaCobrar.Show 1
    End If
  End If
End Sub

Private Sub Form_Load()
  rstCuentasCobrar.CursorLocation = adUseClient
  CboTipo.ListIndex = 0
  Filtrar
End Sub

Sub Filtrar()
  Dim Consulta As String
  Consulta = "SELECT sql_ic_cartera_edades.* from sql_ic_cartera_edades ORDER BY RazonSocial"
  'If TxtBuscamos <> "" Then
  '  Consulta = Consulta + " AND NroDocumento like '%" & TxtBuscamos.Text & "%' order by NroDocumento"
  'End If
  AbrirRecorset rstCuentasCobrar, Consulta, CnnPrincipal, adOpenStatic, adLockReadOnly
  Set Grilla.DataSource = rstCuentasCobrar
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstCuentasCobrar = Nothing
End Sub
