VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmEjecutarScripts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ejecutar scripts"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   7440
      TabIndex        =   10
      Top             =   7800
      Width           =   2415
   End
   Begin MSComctlLib.ListView LstRegistroScripts 
      Height          =   1455
      Left            =   240
      TabIndex        =   9
      Top             =   6240
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Codigo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Version"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.CommandButton CmdEjecutar 
      Caption         =   "Ejecutar script"
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox TxtVersion 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox TxtCodigoControl 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox TxtRuta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   6135
   End
   Begin VB.TextBox TxtSql 
      Height          =   5055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   960
      Width           =   9735
   End
   Begin VB.CommandButton CmdAbrir 
      Caption         =   "Abrir"
      Height          =   255
      Left            =   6840
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Version:"
      Height          =   195
      Left            =   3240
      TabIndex        =   5
      Top             =   600
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Codigo control:"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label LblRuta 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   390
   End
End
Attribute VB_Name = "FrmEjecutarScripts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAbrir_Click()
  Principal.CDExa.Filter = "Archivo de texto |*.txt"
  Principal.CDExa.DialogTitle = "Archivo text"
  Principal.CDExa.ShowOpen
  TxtRuta.Text = Principal.CDExa.FileName
  TxtCodigoControl.Text = Mid(Principal.CDExa.FileTitle, 1, 3)
  TxtVersion.Text = Mid(Principal.CDExa.FileTitle, 10, 6)
  If Principal.CDExa.FileName <> "" Then
    Dim ReadFileName As String
    ReadFileName = TxtRuta.Text
    Open ReadFileName For Input As #1
    ReadFileName = Input$(LOF(1), 1)
    TxtSql.Text = ReadFileName
    Close #1
  End If
  
End Sub

Private Sub CmdEjecutar_Click()
  Dim arraySql() As String
  Dim i As Integer
  Dim NroRegistro As Integer
If TxtSql.Text <> "" Then
  If TxtCodigoControl.Text <> "" Then
    If TxtVersion.Text <> "" Then
      AbrirRecorset rstUniversal, "SELECT registro_scripts.* FROM registro_scripts WHERE CodigoControl = " & Val(TxtCodigoControl.Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
      If rstUniversal.RecordCount > 0 Then
        MsgBox "Este script ya fue ejecutado", vbCritical
      Else
        AbrirRecorset rstUniversal, "SELECT registro_scripts.*  FROM registro_scripts WHERE CodigoControl = " & Val(TxtCodigoControl.Text) - 1, CnnPrincipal, adOpenDynamic, adLockOptimistic
        If rstUniversal.RecordCount > 0 Then
          arraySql = Split(TxtSql.Text, ";")
          NroRegistro = UBound(arraySql) - 1
          For i = LBound(arraySql) To NroRegistro
            AbrirRecorset rstUniversal, arraySql(i), CnnPrincipal, adOpenDynamic, adLockOptimistic
          Next
          AbrirRecorset rstUniversal, "INSERT INTO `bdkl`.`registro_scripts` (`Codigo`, `CodigoControl`, `Version`, `Instalado`) VALUES (NULL, '" & Val(TxtCodigoControl.Text) & "', '" & TxtVersion.Text & "', '1')", CnnPrincipal, adOpenDynamic, adLockOptimistic
          LlenarRegistroScripts
          MsgBox "Script ejecutado con exito", vbExclamation
        Else
          MsgBox "No existe el script anterior numero " & Val(TxtCodigoControl.Text) - 1, vbCritical
        End If
      End If
      CerrarRecorset rstUniversal
    End If
  End If
End If
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub LlenarRegistroScripts()
  LstRegistroScripts.ListItems.Clear
  AbrirRecorset rstUniversal, "SELECT registro_scripts.* FROM registro_scripts ORDER BY CodigoControl DESC", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  Do While rstUniversal.EOF = False
    Set Item = LstRegistroScripts.ListItems.Add(, , rstUniversal!Codigo)
      Item.SubItems(1) = rstUniversal!CodigoControl & ""
      Item.SubItems(2) = rstUniversal!Version & ""
    rstUniversal.MoveNext
  Loop
  CerrarRecorset rstUniversal
End Sub

Private Sub Form_Load()
  LlenarRegistroScripts
End Sub
