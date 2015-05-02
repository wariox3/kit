VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FmrAnalisisDespachoRuta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Analisis de rentabilidad de despachos por ruta"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   14640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdVer 
      Caption         =   "Ver"
      Height          =   255
      Left            =   11400
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Frame FraContenido 
      Enabled         =   0   'False
      Height          =   8295
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   14415
      Begin VB.CommandButton CmdQuitarMarcadas 
         Caption         =   "Quitar marcadas"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   6400
         Width           =   2415
      End
      Begin VB.CommandButton CmdCvDespacho 
         Caption         =   "Colocar en bandeja temporal"
         Height          =   255
         Left            =   9360
         TabIndex        =   25
         Top             =   6400
         Width           =   2415
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   11880
         TabIndex        =   23
         Top             =   3360
         Width           =   2415
      End
      Begin VB.CommandButton CmdAgregarTodos 
         Caption         =   "Agregar todos"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Frame FraDatos 
         Caption         =   "Notas"
         Enabled         =   0   'False
         Height          =   1335
         Left            =   120
         TabIndex        =   19
         Top             =   6720
         Width           =   9615
         Begin VB.TextBox TxtNotas 
            Height          =   975
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   20
            Top             =   240
            Width           =   9375
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   9840
         TabIndex        =   8
         Top             =   6720
         Width           =   4455
         Begin VB.TextBox TxtTund 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   720
            TabIndex        =   13
            Text            =   "0"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox TxtTKR 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   720
            TabIndex        =   12
            Text            =   "0"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox TxtTKV 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   720
            TabIndex        =   11
            Text            =   "0"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox TxtTFlete 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            TabIndex        =   10
            Text            =   "0"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox TxtTManejo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            TabIndex        =   9
            Text            =   "0"
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "T Unid:"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "T KR:"
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Top             =   600
            Width           =   420
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "T KV:"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   960
            Width           =   405
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "T Flete:"
            Height          =   195
            Left            =   2040
            TabIndex        =   15
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "T Manejo:"
            Height          =   195
            Left            =   2040
            TabIndex        =   14
            Top             =   600
            Width           =   720
         End
      End
      Begin VB.CommandButton CmdAnalizar 
         Caption         =   "Analizar"
         Height          =   255
         Left            =   11880
         TabIndex        =   7
         Top             =   6400
         Width           =   2415
      End
      Begin MSComctlLib.ListView LstGuias 
         Height          =   2655
         Left            =   120
         TabIndex        =   22
         Top             =   3720
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   4683
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Guia"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Destino"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Und"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "KR"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "KV"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Flete"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Manejo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Declara"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Object.Tag             =   "Total pagado proporcional"
            Text            =   "TPP"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Object.Tag             =   "Margen de utilidad"
            Text            =   "MU"
            Object.Width           =   1058
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GrillaDatos 
         Height          =   3015
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   5318
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
            DataField       =   "FhEntradaBodega"
            Caption         =   "Fh Entra"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   9226
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "COIng"
            Caption         =   "C Ing"
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
            DataField       =   "CR"
            Caption         =   "C Resp"
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
            DataField       =   "IdCliente"
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
            DataField       =   "Remitente"
            Caption         =   "Remite"
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
            DataField       =   "NmDestinatario"
            Caption         =   "Destinatario"
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
         BeginProperty Column08 
            DataField       =   "Unidades"
            Caption         =   "Unid"
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
            DataField       =   "KilosReales"
            Caption         =   "KR"
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
            DataField       =   "KilosVolumen"
            Caption         =   "KV"
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
         BeginProperty Column11 
            DataField       =   "KilosFacturados"
            Caption         =   "KF"
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
            DataField       =   "VrDeclarado"
            Caption         =   "Declarado"
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
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column12 
               Alignment       =   1
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column13 
               Alignment       =   1
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column14 
               Alignment       =   1
               Locked          =   -1  'True
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3015
      Begin VB.TextBox TxtMargenStd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   32
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TxtFleteStd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   30
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Margen Std:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Flete Std:"
         Height          =   195
         Left            =   315
         TabIndex        =   29
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame FraParametros 
      Height          =   1095
      Left            =   3240
      TabIndex        =   2
      Top             =   0
      Width           =   6855
      Begin VB.ComboBox CboCapacidad 
         Height          =   315
         ItemData        =   "FmrAnalisisDespachoRuta.frx":0000
         Left            =   5400
         List            =   "FmrAnalisisDespachoRuta.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxtFletePagar 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1455
         TabIndex        =   27
         Text            =   "0"
         Top             =   600
         Width           =   2040
      End
      Begin VB.TextBox TxtNmRuta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   5295
      End
      Begin VB.TextBox TxtIdRuta 
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Capacidad:"
         Height          =   255
         Left            =   4440
         TabIndex        =   34
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Flete a Pagar:"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   600
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   390
      End
   End
End
Attribute VB_Name = "FmrAnalisisDespachoRuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstTemp As New ADODB.Recordset

Private Sub CmdAgregar_Click()
  If rstTemp.State = adStateOpen Then
    If rstTemp.RecordCount > 0 Then
      Agregar 1
       USiguiente rstTemp
    Else
      MsgBox "Ho hay guias para esta ruta", vbCritical
    End If
  End If
End Sub

Private Sub CmdAgregarTodos_Click()
  If rstTemp.State = adStateOpen Then
    If rstTemp.RecordCount > 0 Then
      rstTemp.MoveFirst
      Do While rstTemp.EOF = False
        Agregar 2
        rstTemp.MoveNext
      Loop
      rstTemp.MoveFirst
    Else
      MsgBox "Ho hay guias para esta ruta", vbCritical
    End If
  End If
End Sub

Private Sub CmdAnalizar_Click()
Dim Fc As Single, Fp As Single
  If Val(TxtFletePagar.Text) <= 0 Then
    MsgBox "El flete del conductor no puede ser cero (0)", vbCritical
  Else
    For II = 1 To LstGuias.ListItems.Count
      Fc = LstGuias.ListItems(II).SubItems(7)
      LstGuias.ListItems(II).SubItems(10) = Format((((Fc / TxtTFlete) * 100) * Val(TxtFletePagar.Text)) / 100, "#,##0.00;(#,##0.00)")
      Fp = LstGuias.ListItems(II).SubItems(10)
      LstGuias.ListItems(II).SubItems(11) = Format(((Fc - Fp) / Fp) * 100, "#,##0.00;(#,##0.00)")
    Next
  End If
End Sub



Private Sub CmdCvDespacho_Click()
  MsgBox "Estas guias seran colocadas en una bandeja temporal, para que puedan ser importadas desde un despacho", vbInformation
  AbrirRecorset rstUniversal, "DELETE FROM BandejaTemporal", CnnPrincipal, adOpenDynamic, adLockOptimistic
  For II = 1 To LstGuias.ListItems.Count
    AbrirRecorset rstUniversal, "Insert into BandejaTemporal Values(" & LstGuias.ListItems(II).Text & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Next
  MsgBox LstGuias.ListItems.Count & " Guias colocadas con exito", vbInformation
End Sub

Private Sub CmdQuitarMarcadas_Click()
  II = 1
  While II <= LstGuias.ListItems.Count
    If LstGuias.ListItems(II).Checked = True Then
        TxtTund.Text = Format(TxtTund.Text - LstGuias.ListItems(II).SubItems(4), "#,##0.00;(#,##0.00)")
        TxtTKR.Text = Format(TxtTKR.Text - LstGuias.ListItems(II).SubItems(5), "#,##0.00;(#,##0.00)")
        TxtTKV.Text = Format(TxtTKV.Text - LstGuias.ListItems(II).SubItems(6), "#,##0.00;(#,##0.00)")
        TxtTFlete.Text = Format(TxtTFlete.Text - LstGuias.ListItems(II).SubItems(7), "#,##0.00;(#,##0.00)")
        TxtTManejo.Text = Format(TxtTManejo.Text - LstGuias.ListItems(II).SubItems(8), "#,##0.00;(#,##0.00)")
        LstGuias.ListItems.Remove (II)
    Else
     II = II + 1
    End If
  Wend
End Sub

Private Sub CmdVer_Click()
  If Val(TxtIdRuta.Text) <> 0 Then
    If rstTemp.State = adStateOpen Then rstTemp.Close
    rstTemp.Open "SELECT Guias.Guia, Guias.CR, Guias.COIng, Guias.Remitente, Guias.IdCliente, Guias.NmDestinatario, Guias.IdCiuDestino, Ciudades.NmCiudad, Guias.FhEntradaBodega, Guias.IdRuta, Guias.VrDeclarado, Guias.VrFlete, Guias.VrManejo, Guias.Unidades, Guias.KilosFacturados, Guias.KilosReales, Guias.KilosVolumen, Guias.Estado, Guias.IdDespacho FROM Guias INNER JOIN Ciudades ON Guias.IdCiuDestino = Ciudades.IdCiudad Where Estado ='I' and (Guias.IdRuta = " & Val(TxtIdRuta) & ")", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstTemp.RecordCount > 0 Then
      Set GrillaDatos.DataSource = rstTemp
      FraParametros.Enabled = False
      FraContenido.Enabled = True
      Me.Height = 9915
      Me.Top = 800
    End If
  Else
    MsgBox "Debe seleccionar una ruta", vbCritical: TxtIdRuta.SetFocus
  End If
End Sub
Private Sub Form_Load()
  rstTemp.CursorLocation = adUseClient
  CboCapacidad.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rstTemp = Nothing
End Sub

Private Sub TxtIdRuta_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
      Principal.ToolConsultas1.AbrirConsultaGral "IdRuta", "NmRuta", "Rutas", CnnPrincipal
      TxtIdRuta.Text = Principal.ToolConsultas1.DatLo
  End If
End Sub

Private Sub TxtIdRuta_KeyPress(KeyAscii As Integer)
  ValidarEntrada TxtIdRuta, KeyAscii, 1
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtIdRuta_LostFocus()
  If Val(TxtIdRuta.Text) <> 0 Then
    AbrirRecorset rstUniversal, "SELECT IdRuta, NmRuta From Rutas where IdRuta=" & TxtIdRuta, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversal.EOF = False Then
      TxtNmRuta = rstUniversal!NmRuta & ""
      II = 1
    Else
      TxtNmRuta.Text = "": TxtIdRuta.Text = ""
      II = 0
    End If
    CerrarRecorset rstUniversal
    
    If II = 1 Then
      AbrirRecorset rstUniversal, "Select*from Margenes_Rentabilidad_Ruta where IdRuta=" & Val(TxtIdRuta) & " and peso=" & Val(CboCapacidad.Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        TxtNotas.Text = rstUniversal.Fields("Comentarios") & ""
        TxtFleteStd.Text = rstUniversal.Fields("FleteStd")
        TxtMargenStd.Text = rstUniversal.Fields("Margen")
        TxtFletePagar.Text = rstUniversal.Fields("FleteStd")
      Else
        MsgBox "No hay margen de rentabildiad creado en las tablas de margenes de rentabilidad, que son los parametros para calcular los margenes y analizar estos despacho, ingrese los margenes de rentabilidad", vbCritical
        TxtNmRuta.Text = "": TxtIdRuta.Text = ""
        TxtNotas.Text = ""
        TxtFleteStd.Text = ""
        TxtMargenStd.Text = ""
        
      End If
      CerrarRecorset rstUniversal
    End If
  End If
End Sub

Private Sub Agregar(Cri As Byte)
  Set Item = LstGuias.FindItem(rstTemp.Fields("Guia"))
  If Item Is Nothing Then
    If rstTemp.EOF = False Then
      Set Item = LstGuias.ListItems.Add(, , rstTemp.Fields("Guia"))
        Item.SubItems(1) = rstTemp.Fields("FhEntradaBodega")
        Item.SubItems(2) = rstTemp.Fields("IdCliente")
        Item.SubItems(3) = rstTemp.Fields("NmCiudad")
        Item.SubItems(4) = Format(rstTemp.Fields("Unidades"), "#,##0.00;(#,##0.00)")
        Item.SubItems(5) = Format(rstTemp.Fields("KilosReales"), "#,##0.00;(#,##0.00)")
        Item.SubItems(6) = Format(rstTemp.Fields("KilosVolumen"), "#,##0.00;(#,##0.00)")
        Item.SubItems(7) = Format(rstTemp.Fields("VrFlete"), "#,##0.00;(#,##0.00)")
        Item.SubItems(8) = Format(rstTemp.Fields("VrManejo"), "#,##0.00;(#,##0.00)")
        Item.SubItems(9) = Format(rstTemp.Fields("VrDeclarado"), "#,##0.00;(#,##0.00)")
        TxtTund.Text = Format(TxtTund.Text + rstTemp.Fields("Unidades"), "#,##0.00;(#,##0.00)")
        TxtTKR.Text = Format(TxtTKR.Text + rstTemp.Fields("KilosReales"), "#,##0.00;(#,##0.00)")
        TxtTKV.Text = Format(TxtTKV.Text + rstTemp.Fields("KilosVolumen"), "#,##0.00;(#,##0.00)")
        TxtTFlete.Text = Format(TxtTFlete.Text + rstTemp.Fields("VrFlete"), "#,##0.00;(#,##0.00)")
        TxtTManejo.Text = Format(TxtTManejo.Text + rstTemp.Fields("VrManejo"), "#,##0.00;(#,##0.00)")
    End If
  Else
    If Cri = 1 Then
      MsgBox "La guia ya fue agregada", vbCritical
    End If
  End If
End Sub
