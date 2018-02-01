VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmListas 
   Caption         =   "Listas precios..."
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12540
   ControlBox      =   0   'False
   Icon            =   "FrmListas.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   12540
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid GrillaPrecios 
      Height          =   5415
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
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
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "IdCiudad"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "NmCiudadOrigen"
         Caption         =   "Origen"
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
      BeginProperty Column03 
         DataField       =   "IdProducto"
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
      BeginProperty Column04 
         DataField       =   "NmProducto"
         Caption         =   "Producto"
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
         DataField       =   "VrKilo"
         Caption         =   "Vr Kilo"
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
      BeginProperty Column06 
         DataField       =   "VrUnidad"
         Caption         =   "Vr Unidad"
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
      BeginProperty Column07 
         DataField       =   "VrTonelada"
         Caption         =   "Vr Tonelada"
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
      BeginProperty Column08 
         DataField       =   "KTope"
         Caption         =   "K Tope"
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
      BeginProperty Column09 
         DataField       =   "VrKTope"
         Caption         =   "Vr K Tope"
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
         DataField       =   "VrKAdicional"
         Caption         =   "Vr Adiciona"
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
         DataField       =   "Minimos"
         Caption         =   "Min"
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   2069.858
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   1934.929
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   555.024
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraPrecios 
      Height          =   1815
      Left            =   120
      TabIndex        =   13
      Top             =   6000
      Width           =   12285
      Begin VB.CommandButton CmdExportarBufalo 
         Caption         =   "Exportar bufalo"
         Height          =   255
         Left            =   4680
         TabIndex        =   33
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox TxtIdCiudadDestino 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtNmCiudadDestino 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   31
         Top             =   600
         Width           =   4575
      End
      Begin VB.CommandButton CmdImprimirReporte 
         Caption         =   "Imprimir reporte"
         Height          =   255
         Left            =   6600
         TabIndex        =   30
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton CmdCerrar 
         Caption         =   "Cerrar lista de precios"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton CmdEditarRama 
         Caption         =   "Editar lista en rama"
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agregar"
         Height          =   255
         Left            =   9600
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton CmdQuitar 
         Caption         =   "&Quitar"
         Height          =   255
         Left            =   10920
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox TxtNmProducto 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox TxtNmCiudadOrigen 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox TxtIdCiudadOrigen 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtIdProducto 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin MSMask.MaskEdBox TxtVlrKilosTope 
         Height          =   255
         Left            =   10680
         TabIndex        =   7
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtKilosTope 
         Height          =   255
         Left            =   9840
         TabIndex        =   6
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtVlrKiloAdicional 
         Height          =   255
         Left            =   10680
         TabIndex        =   8
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtVlrUnidad 
         Height          =   255
         Left            =   7560
         TabIndex        =   4
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtVlrKilo 
         Height          =   255
         Left            =   7560
         TabIndex        =   3
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtVlrTonelada 
         Height          =   255
         Left            =   7560
         TabIndex        =   5
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtKMinimos 
         Height          =   255
         Left            =   10680
         TabIndex        =   9
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label LBL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   32
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Minimos:"
         Height          =   195
         Index           =   1
         Left            =   9960
         TabIndex        =   29
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vr Ton:"
         Height          =   195
         Left            =   6960
         TabIndex        =   24
         Top             =   960
         Width           =   525
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   120
         X2              =   12120
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   7
         X1              =   120
         X2              =   12120
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "$:"
         Height          =   195
         Left            =   10440
         TabIndex        =   22
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   9240
         TabIndex        =   21
         Top             =   240
         Width           =   465
      End
      Begin VB.Label LBL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Origen:"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   20
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Vr Uni:"
         Height          =   195
         Left            =   7005
         TabIndex        =   19
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vr Kilo:"
         Height          =   195
         Left            =   6990
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin VB.Label LBL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Vr K Adicional:"
         Height          =   195
         Index           =   0
         Left            =   9600
         TabIndex        =   16
         Top             =   600
         Width           =   1035
      End
   End
   Begin VB.Label LblTipo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10560
      TabIndex        =   28
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label LblnmListaPrecios 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   27
      Top             =   120
      Width           =   9495
   End
   Begin VB.Label LblIdLista 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FrmListas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub CmdAgregar_Click()
  If Validacion = True Then
    If LblTipo.Caption = "Base De Datos" Then
      AbrirRecorset rstUniversal, "Select*from listaspreciosciudades where IdListaPrecios=" & Val(LblIdLista.Caption) & " AND IdCiudadOrigen=" & Val(TxtIdCiudadOrigen.Text) & "  AND IdCiudad=" & Val(TxtIdCiudadDestino.Text) & " and Idproducto=" & Val(TxtIdProducto), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Else
    End If
    
    If rstUniversal.EOF = False Then
      If MsgBox("El producto ya existe en la lista. ¿desea modificarlo con los siguientes valores?", vbQuestion + vbYesNo, "¿Aceptar modificacion?") = vbYes Then
        AbrirRecorset rstUniversal, "Delete from listaspreciosciudades where IdListaPrecios= " & Val(LblIdLista.Caption) & " AND IdCiudadOrigen=" & Val(TxtIdCiudadOrigen.Text) & " AND IdCiudad=" & Val(TxtIdCiudadDestino.Text) & " and  idproducto=" & Val(TxtIdProducto.Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
        AbrirRecorset rstUniversal, "INSERT INTO listaspreciosciudades VALUES (" & Val(LblIdLista.Caption) & ", " & Val(TxtIdCiudadOrigen.Text) & "," & Val(TxtIdCiudadDestino.Text) & ", " & Val(TxtIdProducto) & ", " & Val(TxtVlrKilo.Text) & ", " & Val(TxtVlrUnidad.Text) & ", " & Val(TxtVlrTonelada.Text) & ", " & Val(TxtKilosTope.Text) & ", " & Val(TxtVlrKilosTope) & ", " & Val(TxtVlrKiloAdicional.Text) & ", " & Val(TxtKMinimos.Text) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
      Else
        TxtIdCiudadDestino.SetFocus
        Exit Sub
      End If
    Else
      AbrirRecorset rstUniversal, "INSERT INTO listaspreciosciudades VALUES (" & Val(LblIdLista.Caption) & ", " & Val(TxtIdCiudadOrigen.Text) & "," & Val(TxtIdCiudadDestino.Text) & ", " & Val(TxtIdProducto) & ", " & Val(TxtVlrKilo.Text) & ", " & Val(TxtVlrUnidad.Text) & ", " & Val(TxtVlrTonelada.Text) & ", " & Val(TxtKilosTope.Text) & ", " & Val(TxtVlrKilosTope) & ", " & Val(TxtVlrKiloAdicional.Text) & ", " & Val(TxtKMinimos.Text) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
    End If
    Limpiar
    ActualizarGrilla
    TxtIdCiudadDestino.SetFocus
  End If
End Sub

Private Sub CmdCerrar_Click()
  Unload Me
End Sub

Private Sub CmdEditarRama_Click()
  If rstListaPrecios.RecordCount > 0 Then
    If MsgBox("Se va a crear un temporal de esta informacion para asegurar la integridad de la lista de precios" & Chr(13) & "¿Esta seguro de editar esta lista en rama?", vbQuestion + vbYesNo, "Editar lista") = vbYes Then
      AbrirRecorset rstUniversal, "Delete from TemPrecios", CnnPrincipal, adOpenDynamic, adLockOptimistic
      rstListaPrecios.MoveFirst
      IniProg (rstListaPrecios.RecordCount)
      For II = 1 To rstListaPrecios.RecordCount
        AbrirRecorset rstUniversal, "INSERT INTO TemPrecios VALUES (" & II & ", " & rstListaPrecios!IdListaPrecios & ", " & Val(rstListaPrecios!IdCiudad) & ", " & Val(rstListaPrecios!IdProducto) & ", " & Val(rstListaPrecios!VrKilo) & ", " & Val(rstListaPrecios!VrUnidad) & ", " & Val(rstListaPrecios!VrTonelada) & ", " & Val(rstListaPrecios!KTope) & ", " & Val(rstListaPrecios!VrKTope) & ", " & Val(rstListaPrecios!VrKAdicional) & ", " & Val(rstListaPrecios!Minimos) & ", " & rstListaPrecios!IdCiudadOrigen & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
        Prog (II)
        rstListaPrecios.MoveNext
      Next
      FinProg
      FrmEditarRama.Show 1
      rstListaPrecios.Requery
    End If
  Else
    MsgBox "No se puede editar una lista en rama si la lista no tiene registros", vbCritical
  End If
End Sub

Private Sub CmdExportarBufalo_Click()
  Dim RutaSalida As String
  Dim o_Excel     As Object
  Dim o_Libro     As Object
  Dim o_Hoja      As Object
  Dim Fila        As Long
  Dim Columna     As Long
  Dim rstListaPrecioDetalle As New ADODB.Recordset
  rstListaPrecioDetalle.CursorLocation = adUseClient
  
On Error GoTo Error_Handler

  Principal.CDExa.DialogTitle = "Guardar como"
  Principal.CDExa.Filter = "Archivo Excel|*.xls"
  Principal.CDExa.ShowSave
  If Principal.CDExa.FileName <> "" Then
    RutaSalida = Principal.CDExa.FileName
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
    o_Hoja.Cells(1, 1).Value = "empresa"
    o_Hoja.Cells(1, 2).Value = "origen"
    o_Hoja.Cells(1, 3).Value = "destino"
    o_Hoja.Cells(1, 4).Value = "producto"
    o_Hoja.Cells(1, 5).Value = "kilo"
    o_Hoja.Cells(1, 6).Value = "unidad"
    
    FufuSt = "Select listasprecios.codigo_empresa_bufalo, IdCiudadOrigen, IdCiudad, IdProducto, VrKilo, VrUnidad from listaspreciosciudades " & _
    "left join listasprecios ON listaspreciosciudades.IdListaPrecios = listasprecios.IdListaPrecios " & _
    " where listaspreciosciudades.IdListaPrecios=" & Val(LblIdLista.Caption)
    AbrirRecorset rstListaPrecioDetalle, FufuSt, CnnPrincipal, adOpenDynamic, adLockOptimistic
    
    II = 2
    If rstListaPrecioDetalle.RecordCount > 0 Then
      Do While rstListaPrecioDetalle.EOF = False
        o_Hoja.Cells(II, 1).Value = rstListaPrecioDetalle.Fields("codigo_empresa_bufalo")
        o_Hoja.Cells(II, 2).Value = rstListaPrecioDetalle.Fields("IdCiudadOrigen") & ""
        o_Hoja.Cells(II, 3).Value = rstListaPrecioDetalle.Fields("IdCiudad") & ""
        o_Hoja.Cells(II, 4).Value = rstListaPrecioDetalle.Fields("IdProducto") & ""
        o_Hoja.Cells(II, 5).Value = rstListaPrecioDetalle.Fields("VrKilo") & ""
        o_Hoja.Cells(II, 6).Value = rstListaPrecioDetalle.Fields("VrUnidad") & ""
        II = II + 1
        rstListaPrecioDetalle.MoveNext
      Loop
    End If
    o_Libro.Close True, RutaSalida
    o_Excel.Quit
    
  End If
  Exit Sub
Error_Handler:
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
        
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub CmdImprimirReporte_Click()
  Mostrar_Reporte CnnPrincipal, 43, "Select*from sql_ilp_imprimir_lista where IdListaPrecios = " & Val(LblIdLista.Caption), "", 2
End Sub

Private Sub CmdQuitar_Click()
  If rstListaPrecios.State = adStateOpen Then
    If rstListaPrecios.EOF = False Then
      If MsgBox("¿Desea eliminar el producto " & rstListaPrecios.Fields("NmProducto") & " Para la ciudad " & rstListaPrecios.Fields("NmCiudad"), vbYesNo + vbQuestion, "¿Desea eliminar?") = vbYes Then
        AbrirRecorset rstUniversal, "Delete From ListasPreciosCiudades Where IdListaPrecios=" & Val(LblIdLista.Caption) & " AND IdCiudadOrigen=" & Val(rstListaPrecios.Fields("IdCiudadOrigen")) & " AND IdCiudad=" & Val(rstListaPrecios.Fields("IdCiudad")) & " and IdProducto=" & Val(rstListaPrecios.Fields("IdProducto")), CnnPrincipal, adOpenDynamic, adLockOptimistic
        rstListaPrecios.Requery
      End If
    End If
  End If
End Sub

Private Sub Form_Load()
  Dim Coperaciones As String
  Principal.Caption = Principal.Caption & " :: LISTA :: [" & FufuSt & "]"
  Set GrillaPrecios.DataSource = rstListaPrecios
  LblIdLista.Caption = FufuLo
  LblnmListaPrecios.Caption = FufuSt
  If II = 1 Then
    LblTipo.Caption = "Base De Datos"
  ElseIf II = 2 Then
    LblTipo.Caption = "Archivo"
  End If
  Coperaciones = GetSetting("Kit Logistics", "Configuracion", "Coperaciones")
  AbrirRecorset rstUniversal, "SELECT centrosoperaciones.* FROM centrosoperaciones WHERE IDPO = " & Coperaciones, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.RecordCount > 0 Then
    TxtIdCiudadOrigen.Text = rstUniversal.Fields("IdCiudad")
  End If
  CerrarRecorset rstUniversal
End Sub
Private Sub Form_Unload(Cancel As Integer)
  If CnnAcces.State = 1 Then
    CnnAcces.Close
    Set CnnAcces = Nothing
  End If
  If rstListaPrecios.State = adStateOpen Then
    rstListaPrecios.Close
  End If
  
  Principal.Caption = "Editor listas de precios 1.0"
  Principal.MnuArchivo.Enabled = True
End Sub



Private Sub GrillaPrecios_HeadClick(ByVal ColIndex As Integer)
  Select Case ColIndex
    Case 1
      rstListaPrecios.Sort = "NmCiudad Asc"
    Case 3
      rstListaPrecios.Sort = "NmProducto Asc"
  End Select
End Sub

Private Sub TxtIdCiudadDestino_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    FrmConsultaCiudades.Show 1
    TxtIdCiudadDestino = FufuLo
  End If
End Sub
Private Sub TxtIdCiudadDestino_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtIdCiudadDestino, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdCiudadDestino_Validate(Cancel As Boolean)
  If Val(TxtIdCiudadDestino.Text) <> 0 Then
    AbrirRecorset rstUniversal, "SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & TxtIdCiudadDestino, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      TxtNmCiudadDestino = rstUniversal!nmCiudad & ""
    Else
      TxtNmCiudadDestino = "": TxtIdCiudadDestino = ""
    End If
  End If
End Sub

Private Sub TxtIdCiudadOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    FrmConsultaCiudades.Show 1
    TxtIdCiudadOrigen = FufuLo
  End If
End Sub

Private Sub TxtIdCiudadOrigen_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtIdCiudadOrigen, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdCiudadOrigen_Validate(Cancel As Boolean)
  If Val(TxtIdCiudadOrigen.Text) <> 0 Then
    AbrirRecorset rstUniversal, "SELECT IdCiudad, NmCiudad From Ciudades where IdCiudad=" & TxtIdCiudadOrigen, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      TxtNmCiudadOrigen = rstUniversal!nmCiudad & ""
    Else
      TxtNmCiudadOrigen = "": TxtIdCiudadOrigen = ""
    End If
  End If
End Sub

Private Sub TxtIdProducto_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    FrmConsultaProductos.Show 1
    TxtIdProducto = FufuLo
  End If
End Sub
Private Sub TxtIdProducto_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtIdProducto, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdProducto_Validate(Cancel As Boolean)
  If Val(TxtIdProducto.Text) <> 0 Then
    AbrirRecorset rstUniversal, "SELECT IdProducto, NmProducto From Productos where IdProducto=" & TxtIdProducto, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      TxtNmProducto = rstUniversal!nmproducto & ""
    Else
      TxtNmProducto = "": TxtIdProducto = ""
    End If
  End If
End Sub

Private Sub TxtKilosTope_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtKilosTope, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub



Private Sub TxtKMinimos_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtKMinimos, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtVlrKilo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub TxtVlrKiloAdicional_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub
Private Sub TxtVlrKilosTope_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtVlrTonelada_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtVlrUnidad_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Function Validacion() As Boolean
  If Val(TxtIdCiudadDestino.Text) <> 0 Then
    If Val(TxtIdProducto.Text) <> 0 Then
      If Val(TxtVlrKilo.Text) = 0 And Val(TxtVlrUnidad.Text) = 0 And Val(TxtVlrTonelada.Text) = 0 And Val(TxtKilosTope.Text) = 0 And Val(TxtVlrKilosTope.Text) = 0 And Val(TxtVlrKiloAdicional.Text) = 0 Then
        MsgBox "Al menos debe tener un precio": Validacion = False: TxtVlrKilo.SetFocus
      Else
        Validacion = True
      End If
    Else
      MsgBox "El precio debe tener un producto": TxtIdProducto.SetFocus: Validacion = False
    End If
  Else
    MsgBox "El precio debe tener una ciudad": TxtIdCiudadDestino.SetFocus: Validacion = False
  End If
End Function
Sub Limpiar()
  TxtIdCiudadDestino.Text = ""
  TxtNmCiudadDestino.Text = ""
  TxtIdProducto.Text = ""
  TxtNmProducto.Text = ""
  TxtVlrKilo.Text = ""
  TxtVlrUnidad.Text = ""
  TxtVlrTonelada.Text = ""
  TxtKilosTope.Text = ""
  TxtVlrKilosTope.Text = ""
  TxtVlrKiloAdicional.Text = ""
  TxtKMinimos.Text = ""
End Sub
Private Sub ActualizarGrilla()
  rstListaPrecios.Requery
  Set GrillaPrecios.DataSource = rstListaPrecios
End Sub

