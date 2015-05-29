VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmEditarRama 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editar en rama la lista de precios..."
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   12720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "Aplicar llenado"
      Height          =   255
      Left            =   10440
      TabIndex        =   14
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox TxtValor 
      Height          =   285
      Left            =   10920
      TabIndex        =   13
      Top             =   240
      Width           =   1695
   End
   Begin VB.Frame FrMasMen 
      Height          =   885
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   975
      Begin VB.OptionButton OptIgual 
         Caption         =   "Igual"
         Height          =   255
         Left            =   80
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton OptMas 
         Caption         =   "Mas"
         Height          =   255
         Left            =   80
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton OptMenos 
         Caption         =   "Menos"
         Height          =   255
         Left            =   80
         TabIndex        =   7
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame FrOpciones 
      Height          =   885
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox CboOpcLlenado 
         Height          =   315
         ItemData        =   "FrmEditarRama.frx":0000
         Left            =   1680
         List            =   "FrmEditarRama.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   160
         Width           =   2415
      End
      Begin VB.ComboBox CboCampos 
         Height          =   315
         ItemData        =   "FrmEditarRama.frx":0029
         Left            =   1680
         List            =   "FrmEditarRama.frx":0040
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   500
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Campo para aplicar:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   500
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Opcion de llenado:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   160
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdNoGuardarCambios 
      Caption         =   "No guardar cambios"
      Height          =   255
      Left            =   7920
      TabIndex        =   1
      Top             =   8400
      Width           =   2295
   End
   Begin VB.CommandButton CmdGuardarCambios 
      Caption         =   "Guardar cambios"
      Height          =   255
      Left            =   10320
      TabIndex        =   0
      Top             =   8400
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid GrillaPrecios 
      Height          =   7215
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   12726
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "VrKilo"
         Caption         =   "Vr Kilo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "$#,##0;($#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "VrUnidad"
         Caption         =   "Vr Unidad"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "$#,##0;($#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "VrTonelada"
         Caption         =   "Vr Tonelada"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "$#,##0;($#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "KTope"
         Caption         =   "K Tope"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "VrKTope"
         Caption         =   "Vr K Tope"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "$#,##0;($#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "VrKAdicional"
         Caption         =   "Vr Adiciona"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "$#,##0;($#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
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
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1980.284
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1080
         EndProperty
      EndProperty
   End
   Begin VB.Label LblValor 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   195
      Left            =   10440
      TabIndex        =   12
      Top             =   240
      Width           =   405
   End
End
Attribute VB_Name = "FrmEditarRama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstTem As ADODB.Recordset

Private Sub CmdAplicar_Click()
  Dim Valor As Double
  If TxtValor.Text = "" Then TxtValor = 0
  If IsNumeric(TxtValor) = False Then
    MsgBox "Debe ingresar un valor numerico", vbCritical
  End If
  Valor = TxtValor
  If CboOpcLlenado.ListIndex = 0 Then
    If OptMenos.Value = True Then
      IniProg rstUniversal.RecordCount
      For II = 1 To rstUniversal.RecordCount
        AbrirRecorset rstTem, "Update TemPrecios set " & DevCampos & "=" & Val(Val(rstUniversal.Fields(DevCampos)) - ((Val(rstUniversal.Fields(DevCampos)) * Valor) / 100)) & " where ID=" & Val(rstUniversal!ID), CnnPrincipal, adOpenDynamic, adLockOptimistic
        Prog (II)
        rstUniversal.MoveNext
      Next
      FinProg
      MsgBox "Lista acutalizada con exito", vbInformation
    ElseIf OptMas.Value = True Then
      IniProg rstUniversal.RecordCount
      For II = 1 To rstUniversal.RecordCount
        AbrirRecorset rstTem, "Update TemPrecios set " & DevCampos & "=" & Val(Val(rstUniversal.Fields(DevCampos)) + ((Val(rstUniversal.Fields(DevCampos)) * Valor) / 100)) & " where ID=" & Val(rstUniversal!ID), CnnPrincipal, adOpenDynamic, adLockOptimistic
        Prog (II)
        rstUniversal.MoveNext
      Next
      FinProg
      MsgBox "Lista acutalizada con exito", vbInformation
    End If
  ElseIf CboOpcLlenado.ListIndex = 1 Then
    If OptMenos.Value = True Then
      IniProg rstUniversal.RecordCount
      For II = 1 To rstUniversal.RecordCount
        AbrirRecorset rstTem, "Update TemPrecios set " & DevCampos & "=" & Val(Val(rstUniversal.Fields(DevCampos)) - Valor) & " where ID=" & Val(rstUniversal!ID), CnnPrincipal, adOpenDynamic, adLockOptimistic
        Prog (II)
        rstUniversal.MoveNext
      Next
      FinProg
      MsgBox "Lista acutalizada con exito", vbInformation
      
    ElseIf OptMas.Value = True Then
      IniProg rstUniversal.RecordCount
      For II = 1 To rstUniversal.RecordCount
        AbrirRecorset rstTem, "Update TemPrecios set " & DevCampos & "=" & Val(Val(rstUniversal.Fields(DevCampos)) + Valor) & " where ID=" & Val(rstUniversal!ID), CnnPrincipal, adOpenDynamic, adLockOptimistic
        Prog (II)
        rstUniversal.MoveNext
      Next
      FinProg
      MsgBox "Lista acutalizada con exito", vbInformation
      
    ElseIf OptIgual.Value = True Then
      IniProg rstUniversal.RecordCount
      For II = 1 To rstUniversal.RecordCount
        AbrirRecorset rstTem, "Update TemPrecios set " & DevCampos & "=" & Valor & " where ID=" & Val(rstUniversal!ID), CnnPrincipal, adOpenDynamic, adLockOptimistic
        Prog (II)
        rstUniversal.MoveNext
      Next
      FinProg
      MsgBox "Lista acutalizada con exito", vbInformation
    End If
  End If
  AbrirRecorset rstUniversal, "SELECT TemPrecios.*, Ciudades.NmCiudad, Productos.NmProducto FROM (TemPrecios LEFT JOIN Productos ON TemPrecios.IdProducto = Productos.IdProducto) LEFT JOIN Ciudades ON TemPrecios.IdCiudad = Ciudades.IdCiudad", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Set GrillaPrecios.DataSource = rstUniversal
End Sub

Private Sub CmdGuardarCambios_Click()
If MsgBox("¿Esta seguro de actualizar la lista de precios?" & Chr(13) & "- Si presiona en ACEPTAR el sistema actualizara la lista de precios con estos nuevos precios", vbQuestion + vbOKCancel, "¿Desea actualizar los precios?") = vbOK Then
  AbrirRecorset rstTem, "Delete from listaspreciosciudades where IdListaPrecios=" & Val(FrmListas.LblIdLista.Caption), CnnPrincipal, adOpenDynamic, adLockOptimistic
  rstUniversal.MoveFirst
  IniProg (rstUniversal.RecordCount)
  For II = 1 To rstUniversal.RecordCount
    AbrirRecorset rstTem, "INSERT INTO listaspreciosciudades VALUES (" & rstUniversal!IdListaPrecios & ", " & rstUniversal!IdCiudadOrigen & ", " & Val(rstUniversal!IdCiudad) & ", " & Val(rstUniversal!IdProducto) & ", " & Val(rstUniversal!VrKilo) & ", " & Val(rstUniversal!VrUnidad) & ", " & Val(rstUniversal!VrTonelada) & ", " & Val(rstUniversal!KTope) & ", " & Val(rstUniversal!VrKTope) & ", " & Val(rstUniversal!VrKAdicional) & ", " & Val(rstUniversal!Minimos) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
    Prog (II)
    rstUniversal.MoveNext
  Next
  FinProg
  MsgBox "La lista ha sido actualizada con exito", vbInformation
  Unload Me
End If
End Sub

Private Sub CmdNoGuardarCambios_Click()
  CerrarRecorset rstUniversal
  AbrirRecorset rstUniversal, "DELETE From `temprecios`", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Unload Me
End Sub


Private Sub Form_Load()
  AbrirRecorset rstUniversal, "SELECT TemPrecios.*, Ciudades.NmCiudad, Productos.NmProducto FROM (TemPrecios LEFT JOIN Productos ON TemPrecios.IdProducto = Productos.IdProducto) LEFT JOIN Ciudades ON TemPrecios.IdCiudad = Ciudades.IdCiudad", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Set GrillaPrecios.DataSource = rstUniversal
  Set rstTem = New ADODB.Recordset
  rstTem.CursorLocation = adUseClient
End Sub
Private Sub GrillaPrecios_HeadClick(ByVal ColIndex As Integer)
  Select Case ColIndex
    Case 0
      rstUniversal.Sort = "NmCiudad Asc"
    Case 1
      rstUniversal.Sort = "NmProducto Asc"
  End Select
End Sub

Private Sub TxtValor_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtValor, KeyAscii, 2
End Sub

Function DevCampos() As String
Select Case CboCampos.ListIndex
  Case 0
    DevCampos = "VrKilo"
  Case 1
    DevCampos = "VrUnidad"
  Case 2
    DevCampos = "VrTonelada"
  Case 3
    DevCampos = "KTope"
  Case 4
    DevCampos = "VrKTope"
  Case 5
    DevCampos = "VrKAdicional"
End Select
End Function
