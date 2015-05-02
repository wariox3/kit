VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAlertaDocumentosVencidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documentos vencidos"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LstConductores 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4683
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
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Conductor"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "FhLicencia"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "FhSeguridadSocial"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView LstVehiculos 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4683
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Placa"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Vence Soat"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Vence Tecnicomecanica"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Documentos Vencidos de conductores y vehiculos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "FrmAlertaDocumentosVencidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
  Dim rstConductores As New ADODB.Recordset
  rstConductores.CursorLocation = adUseClient
   
  Dim rstVehiculos As New ADODB.Recordset
  rstVehiculos.CursorLocation = adUseClient
   
  LstConductores.ListItems.Clear
  AbrirRecorset rstConductores, "SELECT IdConductor, CONCAT(Nombre, ' ', Apellido1, ' ', Apellido2) AS Nombre, FhVenceLic, FhVenceSeguridadSocial FROM conductores WHERE (FhVenceLic <='" & Format(Date, "yyyy/mm/dd") & "' or FhVenceSeguridadSocial <='" & Format(Date, "yyyy/mm/dd") & "') and ConductorInactivo = 0", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Do While rstConductores.EOF = False
    Set Item = LstConductores.ListItems.Add(, , rstConductores!IdConductor)
    Item.SubItems(1) = rstConductores!Nombre
    Item.SubItems(2) = rstConductores!FhVenceLic
    Item.SubItems(3) = rstConductores!FhVenceSeguridadSocial
    rstConductores.MoveNext
  Loop
  CerrarRecorset rstConductores
  
  LstVehiculos.ListItems.Clear
  AbrirRecorset rstVehiculos, "SELECT IdPlaca, VenceSoat, FhVenceTecnicomecanica FROM vehiculos WHERE (VenceSoat <='" & Format(Date, "yyyy/mm/dd") & "' or FhVenceTecnicomecanica <='" & Format(Date, "yyyy/mm/dd") & "') and Inactivo = 0", CnnPrincipal, adOpenDynamic, adLockOptimistic
  Do While rstVehiculos.EOF = False
    Set Item = LstVehiculos.ListItems.Add(, , rstVehiculos!IdPlaca)
    Item.SubItems(1) = rstVehiculos!VenceSoat
    Item.SubItems(2) = rstVehiculos!FhVenceTecnicomecanica
    rstVehiculos.MoveNext
  Loop
  CerrarRecorset rstVehiculos
  
End Sub
