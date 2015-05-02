VERSION 5.00
Begin VB.Form FrmInfoAsignacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Info asignacion..."
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDatos 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2535
      Begin VB.TextBox TxtKVol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1125
         TabIndex        =   18
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox TxtKReales 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1125
         TabIndex        =   17
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox TxtUnidades 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1125
         TabIndex        =   16
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox TxtRecogidas 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1125
         TabIndex        =   15
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox TxtFlete 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1125
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox TxtFecha 
         Height          =   285
         Left            =   1125
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TxtVehiculo 
         Height          =   285
         Left            =   1125
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TxtAsignacion 
         Height          =   285
         Left            =   1125
         TabIndex        =   11
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Vehiculo:"
         Height          =   195
         Index           =   8
         Left            =   405
         TabIndex        =   10
         Top             =   480
         Width           =   660
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Index           =   7
         Left            =   525
         TabIndex        =   9
         Top             =   840
         Width           =   495
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Flete:"
         Height          =   195
         Index           =   6
         Left            =   675
         TabIndex        =   8
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Recogidas:"
         Height          =   195
         Index           =   5
         Left            =   255
         TabIndex        =   7
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Unidades:"
         Height          =   195
         Index           =   4
         Left            =   345
         TabIndex        =   6
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "K Reales:"
         Height          =   195
         Index           =   3
         Left            =   405
         TabIndex        =   5
         Top             =   2280
         Width           =   690
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "K Vol:"
         Height          =   195
         Index           =   2
         Left            =   645
         TabIndex        =   4
         Top             =   2640
         Width           =   420
      End
      Begin VB.Label LblTitulos 
         AutoSize        =   -1  'True
         Caption         =   "Asignacion:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.CommandButton CmdVerRecogidas 
      Caption         =   "Ver recogidas"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   2295
   End
End
Attribute VB_Name = "FrmInfoAsignacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub CmdVerRecogidas_Click()
  FufuSt = TxtVehiculo
  FufuLo = TxtAsignacion
  II = 0
  FrmVerRecogidas.Show 1
End Sub

Private Sub Form_Load()
  AbrirRecorset rstUniversal, "Select*from VehiculosRecogida where IdAsignacion=" & FufuLo, CnnPrincipal, adOpenDynamic, adLockBatchOptimistic
    TxtAsignacion.Text = rstUniversal!IdAsignacion
    TxtVehiculo.Text = rstUniversal!Placa
    TxtFecha.Text = Format(rstUniversal!Fecha & "", "dd/mm/yyyy")
    TxtFlete.Text = Format(rstUniversal!Flete, "$#,##0;($#,##0)")
    TxtRecogidas.Text = Format(rstUniversal!Rec, "#,##0;(#,##0)")
    TxtUnidades.Text = Format(rstUniversal!Unidades, "#,##0;(#,##0)")
    TxtKReales.Text = Format(rstUniversal!KilosReales, "#,##0;(#,##0)")
    TxtKVol.Text = Format(rstUniversal!KilosVol, "#,##0;(#,##0)")
End Sub
