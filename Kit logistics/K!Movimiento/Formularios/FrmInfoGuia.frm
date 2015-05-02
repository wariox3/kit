VERSION 5.00
Begin VB.Form FrmInfoGuia 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informacion de la guia..."
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdLog 
      Caption         =   "Log"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Frame FraDatosCliente 
      Caption         =   "Datos cliente"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   19
      Top             =   2040
      Width           =   4455
      Begin VB.TextBox TxtTelefono 
         Height          =   285
         Left            =   840
         TabIndex        =   21
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox TxtDireccion 
         Height          =   285
         Left            =   840
         TabIndex        =   20
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Telefono:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.CommandButton CmdVerRelacionesEntregaDoc 
      Caption         =   "Rel Entrega Doc"
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   4455
      Begin VB.TextBox TxtRelEntrega 
         Height          =   285
         Left            =   3360
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TxtGuia 
         Height          =   285
         Left            =   960
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtDespacho 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TxtFactura 
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label LblRelEntrega 
         AutoSize        =   -1  'True
         Caption         =   "Rel Ent:"
         Height          =   195
         Left            =   2640
         TabIndex        =   17
         Top             =   960
         Width           =   570
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Guia:"
         Height          =   195
         Index           =   2
         Left            =   525
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Factura:"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   11
         Top             =   600
         Width           =   585
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Despacho:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   780
      End
   End
   Begin VB.CommandButton CmdVerMonitoreo 
      Caption         =   "Ver Monitoreo"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton CmdInfo 
      Caption         =   "Re-Despachos [Viaje]"
      Height          =   255
      Index           =   7
      Left            =   2280
      TabIndex        =   6
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton CmdInfo 
      Caption         =   "Mercancia"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton CmdInfo 
      Caption         =   "Novedades"
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   4
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton CmdInfo 
      Caption         =   "Factura"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton CmdInfo 
      Caption         =   "Despacho [Viaje]"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton CmdInfo 
      Caption         =   "Cuenta / Cliente"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton CmdInfo 
      Caption         =   "Destino"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Aceptar / Salir"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      Index           =   1
      X1              =   120
      X2              =   4560
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   4560
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "FrmInfoGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdInfo_Click(Index As Integer)

FufuLo = Val(Me.Tag)
Select Case Index
  Case 0
    AbrirRecorset rstUniversal, "Select Guia, IdCiuDestino from guias where Guia=" & Val(Me.Tag), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        FufuLo = rstUniversal!IdCiuDestino
        FrmInfoCiudad.Show 1
      End If
    CerrarRecorset rstUniversal
  Case 1
    AbrirRecorset rstUniversal, "Select Guia, IdDespacho from guias where Guia=" & Val(Me.Tag), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        FufuLo = Val(rstUniversal!IdDespacho & "")
        If FufuLo = 0 Then
          MsgBox "Esta guia no tiene despacho", vbCritical, "Guia sin despacho"
        Else
          FrmInfoDespacho.Show 1
        End If
      End If
    CerrarRecorset rstUniversal
  Case 2
    AbrirRecorset rstUniversal, "Select Guia, IdCliente from guias where Guia=" & Val(Me.Tag), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
      If rstUniversal.EOF = False Then
        FufuLo = Val(rstUniversal!IdCliente)
        FrmInfoCuenta.Show 1
      End If
    CerrarRecorset rstUniversal
  Case 4
    FufuLo = Val(TxtFactura)
    FrmInfoFactura.Show 1
  Case 5
    FufuDo = 2
    FrmNovedades.Show 1
  Case 6
    FrmVerProductos.Show 1
  Case 7
    FufuDo = 2
    FrmVerRedespachos.Show 1
End Select
End Sub

Private Sub CmdLog_Click()
  FufuLo = Val(TxtGuia.Text)
  FrmLogGuias.Show 1
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub CmdVerMonitoreo_Click()
  FufuLo = Val(TxtDespacho.Text)
  FrmVerMonitoreos.Show 1
End Sub

Private Sub CmdVerRelacionesEntregaDoc_Click()
  If Val(TxtRelEntrega.Text) <> 0 Then
    FufuLo = Val(TxtRelEntrega.Text)
    FrmVerRelEntregaDoc.Show 1
  Else
    MsgBox "La guia no tiene relacion de entrega", vbCritical
  End If
End Sub

Private Sub Form_Load()
  Me.Tag = FufuLo
  Dim IdCliente As Double
  AbrirRecorset rstUniversal, "Select Guia, IdDespacho, Cuenta, IdFactura, IdRelEntrega from Guias where Guia=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.EOF = False Then
    TxtGuia.Text = FufuLo
    TxtDespacho.Text = rstUniversal.Fields("IdDespacho") & ""
    TxtFactura.Text = rstUniversal.Fields("IdFactura") & ""
    TxtRelEntrega.Text = rstUniversal.Fields("IdRelEntrega") & ""
    IdCliente = rstUniversal.Fields("Cuenta") & ""
  End If
  CerrarRecorset rstUniversal
  AbrirRecorset rstUniversal, "Select terceros.* from terceros where IdTercero=" & IdCliente, CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.RecordCount > 0 Then
    TxtDireccion.Text = rstUniversal.Fields("Direccion") & ""
    TxtTelefono.Text = rstUniversal.Fields("Telefono") & ""
  End If
  CerrarRecorset rstUniversal
End Sub

