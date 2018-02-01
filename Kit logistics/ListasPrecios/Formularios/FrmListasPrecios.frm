VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmListasPrecios 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listas de precios de la base de datos"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdNueva 
      Caption         =   "Nueva"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "Editar"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton CmdDuplicar 
      Caption         =   "Duplicar"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton CmdVerRefrescar 
      Caption         =   "Ver / Refrescar"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "Eliminar"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   6600
      TabIndex        =   2
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton CmdAbrir 
      Caption         =   "Abrir"
      Default         =   -1  'True
      Height          =   255
      Left            =   6600
      TabIndex        =   1
      Top             =   5760
      Width           =   1455
   End
   Begin MSComctlLib.ListView LstListasPrecios 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9763
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Lista"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Vence"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "E_B"
         Object.Width           =   882
      EndProperty
   End
End
Attribute VB_Name = "FrmListasPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LLenarListasPrecios()
  LstListasPrecios.ListItems.Clear
  AbrirRecorset rstUniversal, "Select*from ListasPrecios Order by NmListaPrecios", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    Do While rstUniversal.EOF = False
      Set Item = LstListasPrecios.ListItems.Add(, , rstUniversal.Fields("IdListaPrecios"))
        Item.SubItems(1) = rstUniversal.Fields("NmListaPrecios") & ""
        Item.SubItems(2) = rstUniversal.Fields("FhVencimiento") & ""
        Item.SubItems(3) = rstUniversal.Fields("codigo_empresa_bufalo") & ""
      rstUniversal.MoveNext
    Loop
  CerrarRecorset rstUniversal
End Sub

Private Sub CmdAbrir_Click()
  Dim strSql As String
  II = 1
  FufuLo = LstListasPrecios.SelectedItem
  FufuSt = LstListasPrecios.ListItems(LstListasPrecios.SelectedItem.Index).SubItems(1)
  strSql = "SELECT listaspreciosciudades.*, ciudades.NmCiudad, productos.NmProducto, ciudades_origen.NmCiudad as NmCiudadOrigen " & _
           "FROM listaspreciosciudades " & _
           "LEFT JOIN ciudades ON listaspreciosciudades.IdCiudad = ciudades.IdCiudad " & _
           "LEFT JOIN productos ON listaspreciosciudades.IdProducto = productos.IdProducto " & _
           "LEFT JOIN ciudades AS ciudades_origen ON listaspreciosciudades.IdCiudadOrigen = ciudades_origen.IdCiudad " & _
           "WHERE IdListaPrecios=" & FufuLo
  AbrirRecorset rstListaPrecios, strSql, CnnPrincipal, adOpenDynamic, adLockOptimistic
  Principal.MnuArchivo.Enabled = False
  Unload Me
  FrmListas.Show
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub CmdDuplicar_Click()
Dim rstListaDetalleOriginal As New ADODB.Recordset
Dim rstListaPrecios As New ADODB.Recordset
rstListaDetalleOriginal.CursorLocation = adUseClient
rstListaPrecios.CursorLocation = adUseClient
AbrirRecorset rstUniversal, "INSERT INTO ListasPrecios (NmListaPrecios, FhVencimiento) VALUES ('" & LstListasPrecios.ListItems(LstListasPrecios.SelectedItem.Index).SubItems(1) & " Copia', '" & Format(Date, "yy-mm-dd") & "')", CnnPrincipal, adOpenDynamic, adLockOptimistic
AbrirRecorset rstUniversal, "SELECT IdListaPrecios FROM listasprecios ORDER BY IdListaPrecios DESC LIMIT 1", CnnPrincipal, adOpenDynamic, adLockOptimistic
FufuLo = rstUniversal!IdListaPrecios
AbrirRecorset rstListaDetalleOriginal, "SELECT listaspreciosciudades.* FROM listaspreciosciudades WHERE IdListaPrecios = " & LstListasPrecios.SelectedItem, CnnPrincipal, adOpenDynamic, adLockOptimistic
While rstListaDetalleOriginal.EOF = False
  AbrirRecorset rstUniversal, "INSERT INTO listaspreciosciudades VALUES (" & FufuLo & ", " & rstListaDetalleOriginal!IdCiudadOrigen & ", " & rstListaDetalleOriginal!IdCiudad & ", " & rstListaDetalleOriginal!IdProducto & ", " & rstListaDetalleOriginal!VrKilo & ", " & rstListaDetalleOriginal!VrUnidad & ", " & rstListaDetalleOriginal!VrTonelada & ", " & rstListaDetalleOriginal!KTope & ", " & rstListaDetalleOriginal!VrKTope & ", " & rstListaDetalleOriginal!VrKAdicional & ", " & rstListaDetalleOriginal!Minimos & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
  rstListaDetalleOriginal.MoveNext
Wend
CerrarRecorset rstListaDetalleOriginal
LLenarListasPrecios
End Sub

Private Sub CmdEditar_Click()
  FufuLo = LstListasPrecios.SelectedItem
  FrmMantenimientoListas.Show 1
  LLenarListasPrecios
End Sub
Private Sub CmdNueva_Click()
  FufuLo = 0
  FrmMantenimientoListas.Show 1
  LLenarListasPrecios
  
End Sub

Private Sub CmdEliminar_Click()
  AbrirRecorset rstUniversal, "Select count(*) as Reg from ListasPreciosCiudades where IdListaPrecios=" & LstListasPrecios.SelectedItem, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If MsgBox("La lista tiene " & rstUniversal.Fields("Reg") & " resgistros de precios " & Chr(13) & "¿Esta seguro de eliminar la lista de precios?", vbYesNo + vbQuestion) = vbYes Then
    AbrirRecorset rstUniversal, "Delete From ListasPreciosCiudades where IdListaPrecios=" & LstListasPrecios.SelectedItem, CnnPrincipal, adOpenDynamic, adLockOptimistic
    AbrirRecorset rstUniversal, "Delete From ListasPrecios where IdListaPrecios=" & LstListasPrecios.SelectedItem, CnnPrincipal, adOpenDynamic, adLockOptimistic
    MsgBox "Listas de precios eliminada con exito", vbInformation
    LLenarListasPrecios
  End If
  CerrarRecorset rstUniversal
End Sub

Private Sub CmdVerRefrescar_Click()
  LLenarListasPrecios
End Sub

Private Sub Form_Load()
  LLenarListasPrecios
End Sub

Private Sub LstListasPrecios_DblClick()
  CmdAbrir_Click
End Sub
