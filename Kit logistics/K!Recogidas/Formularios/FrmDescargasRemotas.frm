VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDescargasRemotas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Descargas remotas..."
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdMarcados 
      Caption         =   "Descargar marcados"
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   6840
      TabIndex        =   0
      Top             =   4200
      Width           =   2175
   End
   Begin MSComctlLib.ListView LstDescargar 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7011
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Asig"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "ID Rec"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Und"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "K Real"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "K Vol"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Hora"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label LblMensaje 
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
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   4095
   End
End
Attribute VB_Name = "FrmDescargasRemotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CnnTemp As New ADODB.Connection
Dim rstTemp As New ADODB.Recordset
Private Sub CmdAceptar_Click()
  Unload Me
End Sub

Private Sub CmdMarcados_Click()
    II = 1
    Do While II <= LstDescargar.ListItems.Count
      If LstDescargar.ListItems(II).Checked = True Then
        AbrirRecorset rstTemp, "Select*from anuncios where IdAnuncio=" & LstDescargar.ListItems(II).SubItems(2) & " and efectiva=1", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstTemp.EOF = True Then
          AbrirRecorset rstUniversal, "Update Anuncios set Efectiva=1, Unidades=" & LstDescargar.ListItems(II).SubItems(3) & ", KilosReales=" & LstDescargar.ListItems(II).SubItems(4) & ", KilosVol=" & LstDescargar.ListItems(II).SubItems(5) & ", TiempoEfectiva='" & LstDescargar.ListItems(II).SubItems(6) & " " & Format(LstDescargar.ListItems(II).SubItems(7), "h:m:s") & "' where IdAnuncio=" & LstDescargar.ListItems(II).SubItems(2), CnnPrincipal, adOpenDynamic, adLockOptimistic
          AbrirRecorset rstUniversal, "Update DatosRecogidas set Descargado=-1 where Id=" & LstDescargar.SelectedItem, CnnTemp, adOpenDynamic, adLockOptimistic
          AbrirRecorset rstUniversal, "Update VehiculosRecogida set Rec=Rec-1 where IdAsignacion=" & LstDescargar.ListItems(II).SubItems(1), CnnPrincipal, adOpenDynamic, adLockOptimistic
        Else
          MsgBox "La recogida " & LstDescargar.ListItems(II).SubItems(2) & " ya fue efectiva"
        End If
        rstTemp.Close
        LstDescargar.ListItems.Remove II
      Else
        II = II + 1
      End If
    Loop
End Sub

Private Sub Form_Load()
On Error GoTo Apertura
  rstTemp.CursorLocation = adUseClient
  CnnTemp.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & GetSetting("Kit Logistics", "Recogidas", "ArchivoInterfaz", "")
  LlenarLista
Apertura:
  If Err.Number = -2147467259 Then
    MsgBox "No se puedo abrir el origen de datos de la informacion"
    LblMensaje.Caption = "No se ha encontrado el origen de datos"
  End If
End Sub
Private Sub LlenarLista()
  If LblMensaje.Caption = "" Then
    AbrirRecorset rstUniversal, "select*from DatosRecogidas where Descargado=0", CnnTemp, adOpenForwardOnly, adLockReadOnly
    Do While Not rstUniversal.EOF
      Set Item = LstDescargar.ListItems.Add(, , rstUniversal!ID)
      Item.SubItems(1) = rstUniversal!IdAsignacion
      Item.SubItems(2) = rstUniversal!IdRecogida
      Item.SubItems(3) = rstUniversal!Unidades
      Item.SubItems(4) = rstUniversal!KilosReales
      Item.SubItems(5) = rstUniversal!KilosVol
      Item.SubItems(6) = rstUniversal!Fecha
      Item.SubItems(7) = rstUniversal!Hora
      rstUniversal.MoveNext
    Loop
  CerrarRecorset rstUniversal
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set CnnTemp = Nothing
  Set rstTemp = Nothing
End Sub
