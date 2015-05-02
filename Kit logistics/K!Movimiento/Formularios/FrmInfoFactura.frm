VERSION 5.00
Begin VB.Form FrmInfoFactura 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informacion de la factura"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6975
      Begin VB.TextBox TxtConceptos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1290
         TabIndex        =   24
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox TxtPlanillas 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1290
         TabIndex        =   23
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox TxtGuias 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1290
         TabIndex        =   22
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox TxtTotalFactura 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1290
         TabIndex        =   21
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox TxtOtros 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1290
         TabIndex        =   19
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox TxtManejo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1290
         TabIndex        =   18
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox TxtFlete 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1290
         TabIndex        =   17
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox TxtNmCliente 
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox TxtIdCliente 
         Height          =   285
         Left            =   1290
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox TxtFechaVence 
         Height          =   285
         Left            =   1290
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TxtFecha 
         Height          =   285
         Left            =   1290
         TabIndex        =   13
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox TxtFactura 
         Height          =   285
         Left            =   1290
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   690
         TabIndex        =   20
         Top             =   2760
         Width           =   405
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Otros:"
         Height          =   195
         Left            =   690
         TabIndex        =   11
         Top             =   3840
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Planillas:"
         Height          =   195
         Left            =   450
         TabIndex        =   10
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Guias:"
         Height          =   195
         Left            =   690
         TabIndex        =   9
         Top             =   3120
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Otros:"
         Height          =   195
         Left            =   690
         TabIndex        =   8
         Top             =   2400
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Menejo:"
         Height          =   195
         Left            =   540
         TabIndex        =   7
         Top             =   2040
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Flete:"
         Height          =   195
         Left            =   720
         TabIndex        =   6
         Top             =   1680
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   585
         TabIndex        =   5
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   615
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha vence:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Factura:"
         Height          =   195
         Left            =   525
         TabIndex        =   2
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   255
      Left            =   5760
      TabIndex        =   0
      Top             =   4440
      Width           =   1335
   End
End
Attribute VB_Name = "FrmInfoFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  AbrirRecorset rstUniversal, "select facturas.*, terceros.RazonSocial From (facturas left join terceros on((facturas.IdCliente = terceros.IDTercero))) Where (facturas.IdFactura =" & FufuLo & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
  If rstUniversal.RecordCount > 0 Then
    TxtFactura = rstUniversal.Fields("IdFactura") & ""
    TxtFecha = Format(rstUniversal.Fields("FhFac") & "", "dd/mm/yyyy")
    TxtFechaVence = Format(rstUniversal.Fields("FhVenceFac") & "", "dd/mm/yyyy")
    TxtIdCliente = rstUniversal.Fields("IdCliente") & ""
    TxtNmCliente = rstUniversal.Fields("RazonSocial") & ""
    TxtFlete = Format(rstUniversal.Fields("TFlete"), "#,##0.00;(#,##0.00)")
    TxtManejo = Format(rstUniversal.Fields("TManejo"), "#,##0.00;(#,##0.00)")
    TxtOtros = Format(rstUniversal.Fields("TOtros"), "#,##0.00;(#,##0.00)")
    TxtTotalFactura = Format(rstUniversal.Fields("TotalFactura"), "#,##0.00;(#,##0.00)")
    TxtGuias = rstUniversal.Fields("NroGuias")
    TxtPlanillas = rstUniversal.Fields("NroPlanillas")
    TxtConceptos = rstUniversal.Fields("NroConceptos")
  End If
  CerrarRecorset rstUniversal
End Sub
