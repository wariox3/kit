VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmImportarGuias 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Importar Guias"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtRuta2 
      Height          =   285
      Left            =   600
      TabIndex        =   13
      Text            =   "C:\"
      Top             =   5760
      Width           =   7575
   End
   Begin VB.CommandButton CmdRuta2 
      Caption         =   "..."
      Height          =   255
      Left            =   8280
      TabIndex        =   12
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox TxtNroFilas 
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Text            =   "1"
      Top             =   6120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Importar excel"
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   6120
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3255
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5741
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdImportarExcel 
      Caption         =   "Abrir excel"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton CmdRuta 
      Caption         =   "..."
      Height          =   255
      Left            =   8280
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar Temporal"
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   255
      Left            =   6960
      TabIndex        =   5
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton CmdImportar 
      Caption         =   "Importar"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   2040
      Width           =   1935
   End
   Begin MSComctlLib.ListView LstGuiasTemporal 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Remitente"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Relacion"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Documento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "NmDestinatario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DirDestinatario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "TelDestinatario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "IdOrigen"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "IdDestino"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Observaciones"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton CmdAbrir 
      Caption         =   "Abrir"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox TxtRuta 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "C:\Documents and Settings\Desarrollo\Escritorio\arch\Estructura importacion.csv"
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   5760
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   390
   End
End
Attribute VB_Name = "FrmImportarGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs As ADODB.Recordset
Private Sub CmdAbrir_Click()
  Dim Remitente As String
  Dim Relacion As String
  Dim Documento As String
  Dim NmDestinatario As String
  Dim DirDestinatario As String
  Dim TelDestinatario As String
  Dim IdOrigen As String
  Dim IdDestino As String
  Dim Observaciones As String
  
  LstGuiasTemporal.ListItems.Clear
  Open TxtRuta.Text For Input As #1
  Do While Not EOF(1)
    Input #1, Remitente, Relacion, Documento, NmDestinatario, DirDestinatario, TelDestinatario, IdOrigen, IdDestino, Observaciones
    Set Item = LstGuiasTemporal.ListItems.Add(, , Remitente)
      Item.SubItems(1) = Relacion
      Item.SubItems(2) = Documento
      Item.SubItems(3) = NmDestinatario
      Item.SubItems(4) = DirDestinatario
      Item.SubItems(5) = TelDestinatario
      Item.SubItems(6) = IdOrigen
      Item.SubItems(7) = IdDestino
      Item.SubItems(8) = Observaciones
  Loop
  Close #1
End Sub

Private Sub CmdImportar_Click()
  Dim i As Integer
  With LstGuiasTemporal.ListItems
  For i = 1 To LstGuiasTemporal.ListItems.Count
    AbrirRecorset rstUniversal, "Insert into guias_imp (Remitente, Relacion, Documento, NmDestinatario, DirDestinatario, TelDestinatario, IdOrigen, IdDestino, Observaciones) values ('" & .Item(i) & "', '" & .Item(i).SubItems(1) & "','" & .Item(i).SubItems(2) & "','" & .Item(i).SubItems(3) & "','" & .Item(i).SubItems(4) & "','" & .Item(i).SubItems(5) & "'," & Val(.Item(i).SubItems(6)) & "," & Val(.Item(i).SubItems(7)) & ",'" & .Item(i).SubItems(8) & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Next
  End With
  MsgBox "Proceso de importacion terminado con exito", vbInformation
End Sub

Private Sub CmdImportarExcel_Click()
    Dim rango As String, hoja As String, ruta As String
  
    ruta = TxtRuta2.Text 'ruta del archivo excel
    rango = "A1:K" & Val(TxtNroFilas.Text) 'Rango de datos (opcional)
    hoja = "Hoja1" 'Nombre de la hoja
  
    Importar_Excel ruta, hoja, rango
End Sub

Private Sub CmdLimpiar_Click()
  AbrirRecorset rstUniversal, "Delete from guias_imp", CnnPrincipal, adOpenDynamic, adLockOptimistic
  MsgBox "Proceso terminado con exito", vbInformation
End Sub

Private Sub CmdRuta_Click()
  Principal.CDExa.ShowOpen
  If Principal.CDExa.FileName <> "" Then
   TxtRuta.Text = Principal.CDExa.FileName
  End If

End Sub

Private Sub CmdRuta2_Click()
  Principal.CDExa.ShowOpen
  If Principal.CDExa.FileName <> "" Then
   TxtRuta2.Text = Principal.CDExa.FileName
  End If
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Public Sub Importar_Excel(Libro As String, hoja As String, Optional rango As String = "")
  
    Dim conexion As ADODB.Connection
    Set conexion = New ADODB.Connection
    conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=" & Libro & _
                  ";Extended Properties=""Excel 8.0;HDR=Yes;"""
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenDynamic
        rs.LockType = adLockOptimistic
    
  
    If rango <> ":" Then
       hoja = hoja & "$" & rango
    End If
      
    rs.Open "SELECT * FROM [" & hoja & "]", conexion, , , adCmdText
      
    ' Mostramos los datos en el datagrid
    Set DataGrid1.DataSource = rs
  
End Sub

Private Sub Command1_Click()
  Do While rs.EOF = False
    AbrirRecorset rstUniversal, "Insert into guias_imp (Remitente, Relacion, Documento, NmDestinatario, DirDestinatario, TelDestinatario, IdOrigen, IdDestino, Observaciones, Declarado) values ('" & rs.Fields(1) & "', '" & rs.Fields(2) & "','" & DevDocSinCeros(rs.Fields(0)) & "','" & rs.Fields(4) & "','" & rs.Fields(5) & "','" & rs.Fields(6) & "'," & Val(rs.Fields(7)) & "," & Val(rs.Fields(8)) & ",'" & rs.Fields(9) & "'," & Val(rs.Fields(10)) & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
    rs.MoveNext
  Loop
  MsgBox "Proceso termino con exito"
End Sub

Private Sub Command2_Click()

End Sub
