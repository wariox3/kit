VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmLlenarRelEntegaDoc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Guias Relacion entrega documentos..."
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAgregarPorDocumente 
      Caption         =   "Agregar por documento"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton CmdQuitar 
      Caption         =   "Quitar"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "Agregar por remision"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   8880
      TabIndex        =   0
      Top             =   6120
      Width           =   1575
   End
   Begin MSComctlLib.ListView LstGuias 
      Height          =   5655
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   9975
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Guia"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Documento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "FhEntrada"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cliente"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Destino"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Unidades"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "KReales"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label LblRelacion 
      Caption         =   "0000"
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
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Relacion:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "FrmLlenarRelEntegaDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
  Unload Me
End Sub

Private Sub CmdAgregar_Click()
  Do While Principal.ToolConsultas1.AbrirDevDatos("Digite el numero de Guia", "Digite el numero de guia para agregarle a la relacion", 3, 0) = True
    AgregarGuia Principal.ToolConsultas1.DatLo
  Loop
End Sub

Private Sub AgregarGuia(Guia As Long)
  Set Item = LstGuias.FindItem(Guia)
  If Item Is Nothing Then
    AbrirRecorset rstUniversalSer, "SELECT Guia, FhEntradaBodega, Cliente, DocCliente, Unidades, KilosReales, Ciudades.NmCiudad From Guias, Ciudades where (Guias.IdCiuDestino=Ciudades.IdCiudad) and Guia =" & Guia, CnnPrincipal, adOpenForwardOnly, adLockReadOnly
    If rstUniversalSer.EOF = False Then
      If MsgBox("Esta seguro de agregar la guia nro " & rstUniversalSer.Fields("Guia") & " con documento Nro " & rstUniversalSer.Fields("DocCliente"), vbQuestion + vbYesNo) = vbYes Then
        AbrirRecorset rstUniversal, "Select guia, IdRelEntrega from guias where Relacionada=1 and Guia=" & Guia, CnnPrincipal, adOpenDynamic, adLockOptimistic
        If rstUniversal.EOF = True Then
            Set Item = LstGuias.ListItems.Add(, , rstUniversalSer!Guia)
              Item.SubItems(1) = rstUniversalSer!DocCliente & ""
              Item.SubItems(2) = Format(rstUniversalSer!FhEntradaBodega, "dd/mm")
              Item.SubItems(3) = rstUniversalSer!Cliente & ""
              Item.SubItems(4) = rstUniversalSer!NmCiudad & ""
              Item.SubItems(5) = rstUniversalSer!Unidades
              Item.SubItems(6) = rstUniversalSer!KilosReales
            CerrarRecorset rstUniversal
            AbrirRecorset rstUniversal, "Update Guias set Relacionada=1, IdRelEntrega=" & Val(LblRelacion.Caption) & " where Guia=" & Guia, CnnPrincipal, adOpenDynamic, adLockOptimistic
        Else
          MsgBox "Esta guia ya fue agregada a una relacion", vbCritical
        End If
      End If
    Else
      MsgBox "La guia Nro. [" & Guia & "] NO existe verifique el numero", vbCritical, "La guia no existe"
    End If
    CerrarRecorset rstUniversalSer
  Else
    MsgBox "Esta guia ya fue agregada a esta relacion", vbInformation, "La guia ya fue agregada"
  End If
End Sub

Private Sub CmdAgregarPorDocumente_Click()
  Dim rstGuiaTemp As New ADODB.Recordset
  rstGuiaTemp.CursorLocation = adUseClient
    Do While Principal.ToolConsultas1.AbrirDevDatos("Documento del cliente", "Digite el numero del documento a buscar", 2, 0) = True
      AbrirRecorset rstGuiaTemp, "select Guia from guias where DocCliente='" & DevDocSinCeros(Principal.ToolConsultas1.DatSt) & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstGuiaTemp.RecordCount > 0 Then
        AgregarGuia rstGuiaTemp.Fields("Guia")
      End If
      CerrarRecorset rstGuiaTemp
    Loop
End Sub

Private Sub CmdQuitar_Click()
II = 1
Do While II <= LstGuias.ListItems.Count
  If LstGuias.ListItems(II).Checked = True Then
      'AbrirRecorset rstUniversal, "Update Guia from relentregadocdet Where IdRelacion=" & Val(LblRelacion.Caption) & " and Guia = " & LstGuias.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
      AbrirRecorset rstUniversal, "Update Guias set Relacionada=0, IdRelEntrega=Null where Guia=" & LstGuias.ListItems(II), CnnPrincipal, adOpenDynamic, adLockOptimistic
      LstGuias.ListItems.Remove (II)
  Else
    II = II + 1
  End If
Loop
End Sub

Private Sub Form_Load()
  LblRelacion.Caption = FufuLo
  AbrirRecorset rstUniversal, "SELECT guias.Guia, guias.IdCiuDestino, guias.DocCliente, guias.FhEntradaBodega, guias.KilosReales, guias.Unidades, guias.Cliente, ciudades.NmCiudad, guias.IdRelEntrega" & _
  " FROM guias LEFT JOIN ciudades ON guias.IdCiuDestino = ciudades.IdCiudad" & _
  " WHERE ((guias.IdRelEntrega=" & Val(LblRelacion.Caption) & "));", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  IniProg 1, rstUniversal.RecordCount
  Do While rstUniversal.EOF = False
    Set Item = LstGuias.ListItems.Add(, , rstUniversal!Guia)
    Item.SubItems(1) = rstUniversal!DocCliente & ""
    Item.SubItems(2) = Format(rstUniversal!FhEntradaBodega, "dd/mm")
    Item.SubItems(3) = rstUniversal!Cliente & ""
    Item.SubItems(4) = rstUniversal!NmCiudad & ""
    Item.SubItems(5) = rstUniversal!Unidades
    Item.SubItems(6) = rstUniversal!KilosReales
    Prog rstUniversal.AbsolutePosition
    rstUniversal.MoveNext
  Loop
  FinProg
  CerrarRecorset rstUniversal
  If FufuSt = "S" Then
    CmdAgregar.Enabled = False
    CmdQuitar.Enabled = False
  End If
End Sub
