VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmPrincipal 
   Caption         =   "K!eci [Editor de coordenadas de impresion]"
   ClientHeight    =   6750
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5430
   Icon            =   "FrmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Frame Frame4 
         Caption         =   "Descripcion"
         Height          =   1935
         Left            =   2640
         TabIndex        =   14
         Top             =   2640
         Width           =   2655
         Begin VB.TextBox TxtDescripcion 
            Height          =   1575
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.CheckBox ChkMostrar 
         Caption         =   "&Imprimir"
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   1440
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tamaño del papel"
         Height          =   975
         Left            =   120
         TabIndex        =   4
         Top             =   5640
         Width           =   3495
         Begin MSComCtl2.UpDown UpDAlto 
            Height          =   255
            Left            =   1800
            TabIndex        =   5
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   393216
            Value           =   5
            BuddyControl    =   "LblAlto"
            BuddyDispid     =   196616
            OrigLeft        =   2640
            OrigTop         =   1560
            OrigRight       =   2880
            OrigBottom      =   1815
            Max             =   1200
            Min             =   5
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65537
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDAncho 
            Height          =   255
            Left            =   1800
            TabIndex        =   6
            Top             =   600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   393216
            Value           =   5
            BuddyControl    =   "LblAncho"
            BuddyDispid     =   196615
            OrigLeft        =   2640
            OrigTop         =   1920
            OrigRight       =   2880
            OrigBottom      =   2175
            Max             =   1200
            Min             =   5
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65537
            Enabled         =   -1  'True
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Twip 567/1 cm"
            Height          =   195
            Index           =   3
            Left            =   2160
            TabIndex        =   12
            Top             =   600
            Width           =   1080
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Twip 567/1 cm"
            Height          =   195
            Index           =   2
            Left            =   2160
            TabIndex        =   11
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label LblAncho 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   720
            TabIndex        =   10
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label LblAlto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   720
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Ancho:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   510
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Alto:"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   7
            Top             =   240
            Width           =   315
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Posicion"
         Height          =   975
         Left            =   3720
         TabIndex        =   1
         Top             =   5640
         Width           =   1575
         Begin VB.OptionButton OptHorizontal 
            Caption         =   "Horizontal"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton OptVertical 
            Caption         =   "Vertical"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin MSComCtl2.UpDown UpDTam 
         Height          =   255
         Left            =   4800
         TabIndex        =   16
         Top             =   1800
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "LblTam"
         BuddyDispid     =   196628
         OrigLeft        =   4815
         OrigTop         =   1800
         OrigRight       =   5055
         OrigBottom      =   2055
         Max             =   254
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDY 
         Height          =   255
         Left            =   4800
         TabIndex        =   17
         Top             =   1080
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "LblY"
         BuddyDispid     =   196629
         OrigLeft        =   4800
         OrigTop         =   1080
         OrigRight       =   5040
         OrigBottom      =   1335
         Max             =   500
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDX 
         Height          =   255
         Left            =   4800
         TabIndex        =   18
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "LblX"
         BuddyDispid     =   196630
         OrigLeft        =   4800
         OrigTop         =   720
         OrigRight       =   5040
         OrigBottom      =   975
         Max             =   500
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin MSComctlLib.ListView LstCampos 
         Height          =   3855
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Campo"
            Object.Width           =   4057
         EndProperty
      End
      Begin MSComCtl2.UpDown UpDLong 
         Height          =   255
         Left            =   4800
         TabIndex        =   20
         Top             =   2160
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "LblLong"
         BuddyDispid     =   196621
         OrigLeft        =   4560
         OrigTop         =   3360
         OrigRight       =   4800
         OrigBottom      =   3615
         Max             =   254
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label LblUniv 
         AutoSize        =   -1  'True
         Caption         =   "Longitud:"
         Height          =   195
         Index           =   2
         Left            =   3000
         TabIndex        =   36
         Top             =   2160
         Width           =   660
      End
      Begin VB.Label LblLong 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3720
         TabIndex        =   35
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "De"
         Height          =   255
         Left            =   840
         TabIndex        =   34
         Top             =   240
         Width           =   255
      End
      Begin VB.Label LblDe 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LblTamañoPapel 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   4800
         Width           =   390
      End
      Begin VB.Label LblPath 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   5040
         Width           =   5175
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   120
         X2              =   5280
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label LblNmCampo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2640
         TabIndex        =   30
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label LblTam 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3720
         TabIndex        =   29
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label LblY 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3720
         TabIndex        =   28
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label LblX 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label LblUniv 
         AutoSize        =   -1  'True
         Caption         =   "Tamaño:"
         Height          =   195
         Index           =   3
         Left            =   3030
         TabIndex        =   26
         Top             =   1800
         Width           =   630
      End
      Begin VB.Label LblUniv 
         AutoSize        =   -1  'True
         Caption         =   "Vertical Y:"
         Height          =   195
         Index           =   1
         Left            =   2940
         TabIndex        =   25
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label LblUniv 
         AutoSize        =   -1  'True
         Caption         =   "Horizontal X:"
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   24
         Top             =   720
         Width           =   900
      End
      Begin VB.Label LNumFich 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ml"
         Height          =   195
         Left            =   5160
         TabIndex        =   22
         Top             =   720
         Width           =   150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ml"
         Height          =   195
         Left            =   5160
         TabIndex        =   21
         Top             =   1080
         Width           =   150
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   120
         X2              =   5280
         Y1              =   4680
         Y2              =   4680
      End
   End
   Begin MSComDlg.CommonDialog CDExa 
      Left            =   240
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu MnuAbrir 
         Caption         =   "Abrir archivo de coordenadas"
         Begin VB.Menu MnuAbrirCoordenada 
            Caption         =   "Factura"
            Index           =   1
         End
         Begin VB.Menu MnuAbrirCoordenada 
            Caption         =   "Remision"
            Index           =   2
         End
         Begin VB.Menu MnuAbrirCoordenada 
            Caption         =   "Manifiesto"
            Index           =   3
         End
         Begin VB.Menu MnuAbrirCoordenada 
            Caption         =   "Recibo de caja"
            Index           =   4
         End
         Begin VB.Menu MnuAbrirCoordenada 
            Caption         =   "Orden de recogida"
            Index           =   5
         End
      End
      Begin VB.Menu MnuGuardar 
         Caption         =   "Guardar"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuCerrarArchivo 
         Caption         =   "Cerrar archivo de coordenadas"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu MnuIndice 
         Caption         =   "Indice"
      End
      Begin VB.Menu MnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAcercade 
         Caption         =   "Acerca de..."
      End
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ArchivoAbierto As Boolean
Dim I As Byte
Dim Item As ListItem
Dim TitArchivo As String
Dim Ori As Byte
Dim Alto As Integer
Dim Ancho As Integer
Dim NRP As Integer
Dim NR As Integer
Dim Campos() As TpCoordenadasImpresion
Private Sub Form_Unload(Cancel As Integer)
  If ArchivoAbierto = True Then
    If MsgBox("¿hay un archivo abierto, desea guardar los cambios?", vbQuestion + vbYesNo, "Guardar cambios") = vbYes Then
      GuardarArchivo
    End If
  End If
End Sub

Private Sub LstCampos_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Guardar
  NRP = Item.Index
  Asignar
  LblDe = LstCampos.SelectedItem.Index
End Sub
Private Sub MnuAbrirCoordenada_Click(Index As Integer)
Select Case Index
  Case 1 'Factura
    CDExa.Filter = "Archivo Coordenadas Factura |*.fsl"
    CDExa.DialogTitle = "Archivo Coordenadas (Factura)"
  Case 2
    CDExa.Filter = "Archivo Coordenadas Remision |*.rmci"
    CDExa.DialogTitle = "Archivo Coordenadas (Remision)"
  Case 3
    CDExa.Filter = "Archivo Coordenadas Manifiesto |*.mfto"
    CDExa.DialogTitle = "Archivo Coordenadas (Manifiesto)"
  Case 4
    CDExa.Filter = "Archivo Coordenadas Recibo de caja |*.rcs"
    CDExa.DialogTitle = "Archivo Coordenadas (Recibo de caja)"
  Case 5
    CDExa.Filter = "Archivo Coordenadas orden de recogida |*.ork"
    CDExa.DialogTitle = "Archivo Coordenadas (Orden de recogida)"
End Select
  CDExa.ShowOpen
  If CpExisteFichero(CDExa.FileName) = True Then
    ArchivoAbierto = True
    FraDatos.Enabled = True
    MnuGuardar.Enabled = True
    MnuCerrarArchivo.Enabled = True
    MnuAbrir.Enabled = False
    FraDatos.Visible = True
    LblPath.Caption = CDExa.FileName
    Open LblPath.Caption For Input As #1
    Input #1, TitArchivo, NR, Ori, Alto, Ancho
      If Ori = 1 Then OptVertical = True Else OptHorizontal = True
      LblAlto = Alto
      LblAncho = Ancho
      ReDim Campos(1 To NR)
      For I = 1 To NR
        Input #1, Campos(I).Campo, Campos(I).VX, Campos(I).VY, Campos(I).Mostrar, Campos(I).Tamaño, Campos(I).Descripcion, Campos(I).Longitud
      Next
      LNumFich.Caption = Str(NR) + " Campos"
    Close #1

    For I = 1 To NR
      Set Item = LstCampos.ListItems.Add(, , Campos(I).Campo)
    Next
    NRP = 1
    Asignar
  End If
End Sub
Sub Asignar()
  LblNmCampo = Campos(NRP).Campo
  LblX = Campos(NRP).VX
  LblY = Campos(NRP).VY
  ChkMostrar.Value = Campos(NRP).Mostrar
  LblTam = Campos(NRP).Tamaño
  TxtDescripcion = Campos(NRP).Descripcion
  LblLong = Campos(NRP).Longitud
End Sub
Sub Guardar()
  Campos(NRP).VX = Val(LblX.Caption)
  Campos(NRP).VY = Val(LblY.Caption)
  Campos(NRP).Mostrar = ChkMostrar.Value
  Campos(NRP).Tamaño = Val(LblTam.Caption)
  Campos(NRP).Descripcion = TxtDescripcion
  Campos(NRP).Longitud = Val(LblLong.Caption)
End Sub


Private Sub MnuCerrarArchivo_Click()
  If MsgBox("¿Desea guardar los cambios efectuados?", vbQuestion + vbYesNo, "Guardar cambios") = vbYes Then
    GuardarArchivo
  End If
  ArchivoAbierto = False
  FraDatos.Enabled = False
  MnuGuardar.Enabled = True
  MnuCerrarArchivo.Enabled = True
  MnuAbrir.Enabled = True
  FraDatos.Visible = False
  LblDe.Caption = ""
  LNumFich.Caption = ""
  LblNmCampo.Caption = ""
  LblX.Caption = ""
  LblY.Caption = ""
  LblTam.Caption = ""
  LblLong.Caption = ""
  TxtDescripcion.Text = ""
  LblPath.Caption = ""
  LblAlto.Caption = ""
  LblAncho.Caption = ""
  LstCampos.ListItems.Clear
End Sub

Private Sub MnuGuardar_Click()
  GuardarArchivo
End Sub

Private Sub MnuSalir_Click()
  Unload Me
End Sub
Sub GuardarArchivo()
  Guardar
  Open LblPath.Caption For Output As #1
    If OptVertical.Value = True Then Ori = 1 Else Ori = 2
    Alto = Val(LblAlto)
    Ancho = Val(LblAncho)
  Write #1, TitArchivo, NR, Ori, Alto, Ancho
  For I = 1 To NR
    Write #1, Campos(I).Campo, Campos(I).VX, Campos(I).VY, Campos(I).Mostrar, Campos(I).Tamaño, Campos(I).Descripcion, Campos(I).Longitud
  Next
  Close #1
End Sub
