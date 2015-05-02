VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmVehiculos 
   Caption         =   "Vehiculos..."
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   720
   End
   Begin VB.CommandButton CmdVer 
      Caption         =   "Ver"
      Height          =   255
      Left            =   11400
      TabIndex        =   72
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox PicOrigen 
      Height          =   360
      Left            =   8160
      ScaleHeight     =   300
      ScaleWidth      =   585
      TabIndex        =   71
      Top             =   6720
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   8880
      TabIndex        =   69
      Top             =   5280
      Width           =   2415
      Begin VB.PictureBox PicVehiculo 
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1275
         ScaleWidth      =   2115
         TabIndex        =   70
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame FraFhIngreso 
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      TabIndex        =   66
      Top             =   6720
      Width           =   3015
      Begin VB.TextBox TxtFhIngreso 
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ingreso:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   68
         Top             =   240
         Width           =   570
      End
   End
   Begin VB.Frame FraPropietario 
      Caption         =   "Propietario"
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      TabIndex        =   59
      Top             =   6000
      Width           =   8175
      Begin VB.TextBox TxtConsulta 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   60
         Top             =   240
         Width           =   6135
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   16
         Left            =   360
         TabIndex        =   29
         Tag             =   "1"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         Height          =   195
         Index           =   27
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   210
      End
   End
   Begin VB.Frame FraTenedor 
      Caption         =   "Tenedor"
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      TabIndex        =   56
      Top             =   5280
      Width           =   8175
      Begin VB.TextBox TxtConsulta 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   57
         Top             =   240
         Width           =   6135
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   15
         Left            =   360
         TabIndex        =   28
         Tag             =   "1"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         Height          =   195
         Index           =   19
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   210
      End
   End
   Begin VB.Frame FraAseguradora 
      Caption         =   "Aseguradora"
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      TabIndex        =   51
      Top             =   4560
      Width           =   10695
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   13
         Left            =   360
         MaxLength       =   11
         TabIndex        =   25
         Tag             =   "1"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtConsulta 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   1800
         TabIndex        =   52
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   14
         Left            =   6600
         MaxLength       =   20
         TabIndex        =   26
         Tag             =   "1"
         Top             =   240
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPFehVenPol 
         Height          =   300
         Left            =   9000
         TabIndex        =   27
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16842753
         CurrentDate     =   38030
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Nit"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Pol/SOAT:"
         Height          =   195
         Index           =   12
         Left            =   5760
         TabIndex        =   54
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Vence:"
         Height          =   195
         Index           =   13
         Left            =   8400
         TabIndex        =   53
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Frame FraDatos 
      Enabled         =   0   'False
      Height          =   3855
      Left            =   600
      TabIndex        =   30
      Top             =   600
      Width           =   10695
      Begin VB.CheckBox ChkPropio 
         Alignment       =   1  'Right Justify
         Caption         =   "Vehiculo propio"
         Height          =   255
         Left            =   4560
         TabIndex        =   79
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   26
         Left            =   8040
         MaxLength       =   50
         TabIndex        =   24
         Top             =   3480
         Width           =   2535
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   25
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   23
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   24
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   22
         Top             =   3480
         Width           =   2055
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   23
         Left            =   8760
         MaxLength       =   50
         TabIndex        =   19
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   22
         Left            =   8760
         MaxLength       =   4
         TabIndex        =   18
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox TxtNmCarroceria 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8160
         TabIndex        =   65
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   21
         Left            =   7440
         TabIndex        =   17
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   20
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TxtNmColor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8160
         TabIndex        =   64
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   19
         Left            =   7440
         TabIndex        =   16
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox TxtNmMarca 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8160
         TabIndex        =   63
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   18
         Left            =   7440
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   0
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   0
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   1
         Tag             =   "1"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   2
         Tag             =   "1"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "1"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   8
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   7
         Tag             =   "1"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   17
         Left            =   1200
         MaxLength       =   500
         TabIndex        =   21
         Top             =   3120
         Width           =   9375
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   9
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   8
         Tag             =   "1"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   10
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   9
         Tag             =   "1"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   6
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   11
         Left            =   4800
         MaxLength       =   12
         TabIndex        =   10
         Tag             =   "1"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox ChkRevFisMec 
         Alignment       =   1  'Right Justify
         Caption         =   "Revision tecno mecánica"
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox TxtCampos 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   12
         Left            =   4800
         MaxLength       =   20
         TabIndex        =   11
         Tag             =   "1"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CheckBox ChkInactivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Inactivo"
         Height          =   255
         Left            =   5280
         TabIndex        =   13
         Top             =   2760
         Width           =   975
      End
      Begin MSDataListLib.DataCombo CboLineas 
         Height          =   315
         Left            =   7440
         TabIndex        =   15
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker DTPVenceTecnicomecanica 
         Height          =   300
         Left            =   8760
         TabIndex        =   20
         Top             =   2760
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16842753
         CurrentDate     =   38030
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pagina satelital:"
         Height          =   195
         Left            =   6840
         TabIndex        =   78
         Top             =   3480
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Clave satelital:"
         Height          =   195
         Left            =   3720
         TabIndex        =   77
         Top             =   3480
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Usu satelital:"
         Height          =   195
         Left            =   120
         TabIndex        =   76
         Top             =   3480
         Width           =   900
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Vence tecnicomecanica:"
         Height          =   195
         Index           =   22
         Left            =   6960
         TabIndex        =   75
         Top             =   2760
         Width           =   1770
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Nro. tecnico mecanica:"
         Height          =   195
         Index           =   21
         Left            =   7080
         TabIndex        =   74
         Top             =   2400
         Width           =   1650
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Configuracion:"
         Height          =   195
         Index           =   20
         Left            =   7680
         TabIndex        =   73
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Placa:"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   49
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "P/Remolque:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   48
         Top             =   600
         Width           =   945
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   47
         Top             =   960
         Width           =   570
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Mod. Repot:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   46
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Nro Ejes:"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   45
         Top             =   2040
         Width           =   645
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Peso vacio:"
         Height          =   195
         Index           =   6
         Left            =   3840
         TabIndex        =   44
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Marca:"
         Height          =   195
         Index           =   7
         Left            =   6825
         TabIndex        =   43
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Color:"
         Height          =   195
         Index           =   8
         Left            =   6915
         TabIndex        =   42
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Motor:"
         Height          =   195
         Index           =   9
         Left            =   600
         TabIndex        =   41
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Chasis:"
         Height          =   195
         Index           =   10
         Left            =   600
         TabIndex        =   40
         Top             =   2400
         Width           =   510
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Carroceria:"
         Height          =   195
         Index           =   11
         Left            =   6555
         TabIndex        =   39
         Top             =   1680
         Width           =   765
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Comentarios:"
         Height          =   195
         Index           =   18
         Left            =   240
         TabIndex        =   38
         Top             =   3120
         Width           =   915
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Capacidad Kilos:"
         Height          =   195
         Index           =   14
         Left            =   3495
         TabIndex        =   37
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Capacidad Vol:"
         Height          =   195
         Index           =   15
         Left            =   3600
         TabIndex        =   36
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Index           =   16
         Left            =   720
         TabIndex        =   35
         Top             =   2760
         Width           =   405
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Celular:"
         Height          =   195
         Index           =   17
         Left            =   4200
         TabIndex        =   34
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Reg Nal Carga:"
         Height          =   195
         Index           =   1
         Left            =   3600
         TabIndex        =   33
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Linea:"
         Height          =   195
         Index           =   2
         Left            =   6885
         TabIndex        =   32
         Top             =   600
         Width           =   435
      End
   End
   Begin MSComctlLib.Toolbar ToolVehiculos 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   62
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1005
      ButtonWidth     =   847
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuev"
            Object.ToolTipText     =   "Crear nuevo registro [F9]"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Guar"
            Object.ToolTipText     =   "Guarda la informacio [F10]"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Object.ToolTipText     =   "Editar la informacion guardada [F11]"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Elim"
            Object.ToolTipText     =   "Elimina o anula el registro [F3]"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Can"
            Object.ToolTipText     =   "Cancela la creacion del registro [F4]"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Bus"
            Object.ToolTipText     =   "Buscar [Inicio]"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pri"
            Object.ToolTipText     =   "Ir al primer registro [F5]"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ant"
            Object.ToolTipText     =   "Ir al anterior registro [F6]"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sig"
            Object.ToolTipText     =   "Ir al siguiente registro [F7]"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ult"
            Object.ToolTipText     =   "Ir al ultimo registro [F8]"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cer"
            Object.ToolTipText     =   "Cerrar esta ventana [F12]"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Act"
            Object.ToolTipText     =   "Actualizar la informacion"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imp"
            Object.ToolTipText     =   "Imprimir registro [Fin]"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Car"
            Object.ToolTipText     =   "Carga informacion adicional [Pausa]"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmVehiculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Editando As Boolean
Dim rstVehiculos As New ADODB.Recordset
Dim strSqlVehiculos As String

Private Sub CboLineas_Click(Area As Integer)
  Dim rstLinea As New ADODB.Recordset
  rstLinea.CursorLocation = adUseClient
  
  AbrirRecorset rstLinea, "Select IdLinea from lineas where NmLinea='" & CboLineas & "' AND IdMarca = " & Val(TxtCampos(18).Text), CnnPrincipal, adOpenForwardOnly, adLockReadOnly
  If rstLinea.EOF = False Then
    CboLineas.Tag = rstLinea!IdLinea
    TxtCampos(20).Text = rstLinea!IdLinea
  Else
    CboLineas.Tag = "": TxtCampos(20).Text = ""
  End If
  CerrarRecorset rstLinea
End Sub

Private Sub ChkInactivo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub ChkRevFisMec_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub CmdBorrar_Click()
  PicVehiculo.Picture = Nothing
  PicVehiculo.ToolTipText = ""
End Sub

Private Sub CmdCargar_Click()
On Error GoTo errSub
  
  With Principal.CommonDialog1
       
       .DialogTitle = "Seleccionar un archivo"
       .Filter = "Archivos Todos|*.*|BMP|*.bmp|Archivos JPG|*.jpg"
       
       .ShowOpen
       
       If .FileName = "" Then Exit Sub
       PicOrigen = LoadPicture(.FileName)
       PicVehiculo.PaintPicture PicOrigen, 0, 0, PicVehiculo.ScaleWidth, PicVehiculo.ScaleHeight
       PicVehiculo.ToolTipText = .FileName
  End With
  
  Exit Sub
  
errSub:
    
  If Err.Number = 53 Then
     MsgBox "No se puede cargar dicho archivo, verifique la ruta", vbCritical
  End If
End Sub

Private Sub CmdVer_Click()
    FufuSt = TxtCampos(0).Text
    FrmVerImagen.Show 1
End Sub



Private Sub DTPFehVenPol_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub DTPFhVenceTarjOp_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub DTPVenceTecnicomecanica_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  LaTecla (KeyCode), ToolVehiculos
End Sub
Private Sub Form_Load()
  rstVehiculos.CursorLocation = adUseServer
  strSqlVehiculos = "SELECT vehiculos.*, " & _
                    "vehiculos.Inactivo, " & _
                    "marcas.NmMarca, " & _
                    "lineas.NmLinea, " & _
                    "colores.NmColor, " & _
                    "carrocerias.NmCarroceria, " & _
                    "Aseguradora.RazonSocial AS NmAseguradora, " & _
                    "Tenedor.RazonSocial AS NmTenedor, " & _
                    "Propietario.RazonSocial AS NmPropietario " & _
                    "FROM vehiculos " & _
                    "LEFT JOIN marcas ON vehiculos.IdMarca = marcas.IdMarca " & _
                    "LEFT JOIN lineas ON vehiculos.IdLinea = lineas.IdLinea " & _
                    "LEFT JOIN carrocerias ON vehiculos.IdCarroceria = carrocerias.IdCarroceria " & _
                    "LEFT JOIN colores ON vehiculos.IdColor = colores.IdColor " & _
                    "LEFT JOIN terceros AS Aseguradora ON vehiculos.IdAseguradora = Aseguradora.IDTercero " & _
                    "LEFT JOIN terceros AS Tenedor ON vehiculos.IdTenedor = Tenedor.IDTercero " & _
                    "LEFT JOIN terceros AS Propietario ON vehiculos.IdPropietario = Propietario.IDTercero "
                    
  AbrirRecorset rstVehiculos, strSqlVehiculos & " WHERE vehiculos.Inactivo = 0", CnnPrincipal, adOpenDynamic, adLockOptimistic
  IconosTool ToolVehiculos, Principal.IgListTool
  Asignar rstVehiculos
End Sub
Private Sub Asignar(rstAsignar As ADODB.Recordset)
  For II = 0 To 26
    TxtCampos(II).Text = rstAsignar.Fields(II) & ""
  Next
  DTPFehVenPol = rstAsignar.Fields("VenceSoat")
  DTPVenceTecnicomecanica = rstAsignar.Fields("FhVenceTecnicomecanica")
  ChkRevFisMec.Value = rstAsignar.Fields("RevFisicoMec")
  ChkInactivo.Value = rstAsignar.Fields("Inactivo")
  ChkPropio.Value = rstAsignar.Fields("VehiculoPropio")
  TxtFhIngreso = rstAsignar.Fields("FhIngreso")
  
  If CpExisteFichero(GetSetting("Kit Logistics", "Configuracion", "DirImagenes") & "vehiculos\" & TxtCampos(0).Text & ".jpg") = True Then
    PicOrigen = LoadPicture(GetSetting("Kit Logistics", "Configuracion", "DirImagenes") & "vehiculos\" & TxtCampos(0).Text & ".jpg")
    PicVehiculo.PaintPicture PicOrigen, 0, 0, PicVehiculo.ScaleWidth, PicVehiculo.ScaleHeight
  Else
    PicVehiculo.Picture = Nothing
  End If
  TxtNmMarca.Text = rstAsignar.Fields("NmMarca") & ""
  CboLineas.Text = rstAsignar.Fields("NmLinea") & ""
  TxtNmColor.Text = rstAsignar.Fields("NmColor") & ""
  TxtNmCarroceria.Text = rstAsignar.Fields("NmCarroceria") & ""
  TxtConsulta(13).Text = rstAsignar.Fields("NmAseguradora") & ""
  TxtConsulta(15).Text = rstAsignar.Fields("NmTenedor") & ""
  TxtConsulta(16).Text = rstAsignar.Fields("NmPropietario") & ""
End Sub
Sub limpiar()
  For II = 0 To 26
    TxtCampos(II).Text = ""
  Next
  DTPFehVenPol = Date
  DTPVenceTecnicomecanica = Date
  ChkRevFisMec.Value = 0
  ChkInactivo.Value = 0
  CboLineas.Text = ""
  LimpiarConsulta
End Sub
Sub LimpiarConsulta()
    TxtNmMarca.Text = ""
  TxtNmColor.Text = ""
  CboLineas.Text = ""
  TxtNmCarroceria.Text = ""
  TxtConsulta(13).Text = ""
  TxtConsulta(15).Text = ""
  TxtConsulta(16).Text = ""
End Sub
Sub Desbloquear()
  FraDatos.Enabled = True
    If CpPermisoEspecial(9, CodUsuarioActivo, CnnPrincipal) = False Then
    ChkInactivo.Enabled = False
  End If
  FraAseguradora.Enabled = True
  FraTenedor.Enabled = True
  FraPropietario.Enabled = True
  BotTool 3, 17, ToolVehiculos, True
  
End Sub
Sub Bloquear()
  ChkInactivo.Enabled = True
  FraDatos.Enabled = False
  FraAseguradora.Enabled = False
  FraTenedor.Enabled = False
  FraPropietario.Enabled = False
  BotTool 3, 17, ToolVehiculos, False
End Sub
Sub AccionTool(Indice As Byte)
  Select Case Indice
    Case 3  'Nuevo
      If CpPermiso(6, CodUsuarioActivo, 2, CnnPrincipal) = True Then
        If Principal.ToolConsultas1.AbrirDevDatos("Digite la Placa", "Digite la placla para el nuevo vehiculo", 4, 0) = True Then
          If ExRecorset("Select IdPlaca from vehiculos where IdPlaca='" & FufuSt & "'") = False Then
            Desbloquear
            limpiar
            TxtFhIngreso = Date
            TxtCampos(0) = UCase(Principal.ToolConsultas1.DatSt)
            TxtCampos(1).SetFocus
          Else
            MsgBox "Este vehiculo ya existe", vbCritical
          End If
        End If
      End If
    Case 4  'Guardar
      If Editando = True Then
        If MsgBox("¿Desea aceptar la modificacion?", vbYesNo + vbQuestion, "Modificar registro") = vbYes Then
          If Validacion = True Then
            Bloquear
            MsgBox PicVehiculo.ToolTipText
            MsgBox DvRutaEsp(PicVehiculo.ToolTipText)
            
            AbrirRecorset rstUniversal, "Update Vehiculos set PlacaRemolque='" & TxtCampos(1).Text & "', Modelo='" & TxtCampos(2).Text & "', ModeloRep='" & TxtCampos(3).Text & "', Motor='" & TxtCampos(4).Text & "', NroEjes='" & TxtCampos(5).Text & "', Chasis='" & TxtCampos(6).Text & "', Serie='" & TxtCampos(7).Text & "', PesoVacio='" & TxtCampos(8).Text & "', CapKilos='" & TxtCampos(9).Text & "', CapVol='" & TxtCampos(10).Text & "', Cel='" & TxtCampos(11).Text & "', RegNalCarga='" & TxtCampos(12).Text & "', IdAseguradora='" & TxtCampos(13).Text & "', Soat='" & TxtCampos(14).Text & "', IdTenedor='" & TxtCampos(15).Text & "', IdPropietario='" & TxtCampos(16).Text & "', Comentarios='" & TxtCampos(17).Text & "', IdMarca=" & Val(TxtCampos(18).Text) & ", " & _
            " IdColor=" & Val(TxtCampos(19).Text) & ", IdLinea=" & Val(TxtCampos(20).Text) & ", IdCarroceria=" & Val(TxtCampos(21).Text) & ", VehConfiguracion='" & TxtCampos(22).Text & "', NumeroTecnicomecanica='" & TxtCampos(23).Text & "', FhVenceTecnicomecanica='" & Format(DTPVenceTecnicomecanica.Value, "yyyy/mm/dd") & "', VenceSoat='" & Format(DTPFehVenPol.Value, "yyyy/mm/dd") & "', RevFisicoMec=" & ChkRevFisMec & ", Inactivo=" & ChkInactivo & ", VehiculoPropio=" & ChkPropio & ", UsuarioSatelital='" & TxtCampos(24).Text & "', ClaveSatelital='" & TxtCampos(25).Text & "', PaginaSatelital='" & TxtCampos(26).Text & "' where IdPlaca='" & TxtCampos(0).Text & "'", CnnPrincipal, adOpenDynamic, adLockOptimistic
            Editando = False
            AccionTool 17
          End If
        Else
          Editando = False
          AccionTool 7
        End If
      Else
        If Validacion = True Then
          AbrirRecorset rstUniversal, "INSERT INTO Vehiculos (IdPlaca, PlacaRemolque, Modelo, ModeloRep, Motor, NroEjes, Chasis, Serie, PesoVacio, CapKilos, CapVol, Cel, RegNalCarga, IdAseguradora, Soat, IdTenedor, IdPropietario, Comentarios, IdMarca, IdColor, IdLinea, IdCarroceria, VehConfiguracion, NumeroTecnicomecanica, FhVenceTecnicomecanica, VenceSoat, RevFisicoMec, Inactivo, FhIngreso, UsuarioSatelital, ClaveSatelital, PaginaSatelital, VehiculoPropio) " & _
          " VALUES ('" & TxtCampos(0).Text & "', '" & TxtCampos(1).Text & "', '" & TxtCampos(2).Text & "', '" & TxtCampos(3).Text & "', '" & TxtCampos(4).Text & "', '" & TxtCampos(5).Text & "', '" & TxtCampos(6).Text & "', '" & TxtCampos(7).Text & "', '" & TxtCampos(8).Text & "', '" & TxtCampos(9).Text & "', '" & TxtCampos(10).Text & "', '" & TxtCampos(11).Text & "', '" & TxtCampos(12).Text & "', '" & TxtCampos(13).Text & "', '" & TxtCampos(14).Text & "', '" & TxtCampos(15).Text & "', '" & TxtCampos(16).Text & "', '" & TxtCampos(17).Text & "', " & Val(TxtCampos(18).Text) & ", " & Val(TxtCampos(19).Text) & ", " & Val(TxtCampos(20).Text) & ", " & Val(TxtCampos(21).Text) & ",'" & TxtCampos(22).Text & "' , '" & TxtCampos(23).Text & "', '" & Format(DTPVenceTecnicomecanica.Value, "yyyy/mm/dd") & "', '" & Format(DTPFehVenPol.Value, "yyyy/mm/dd") & "', " & ChkRevFisMec.Value & ", " & ChkInactivo.Value & ", '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "'," & _
          "'" & TxtCampos(24).Text & "','" & TxtCampos(25).Text & "','" & TxtCampos(26).Text & "', " & ChkPropio.Value & ")", CnnPrincipal, adOpenDynamic, adLockOptimistic
          Bloquear
        End If
      End If
    Case 5  'Editar
      If CpPermiso(6, CodUsuarioActivo, 3, CnnPrincipal) = True Then
        Editando = True
        Desbloquear
      End If
    Case 6 'Eliminar
      If CpPermiso(6, CodUsuarioActivo, 4, CnnPrincipal) = True Then
        MsgBox "No se pueden eliminar vehiculos", vbCritical
      End If
    Case 7 'Cancelar
      If Editando = True Then
        AccionTool 4
      Else
        Asignar rstVehiculos
        Bloquear
      End If
    Case 9  'Buscar
      Principal.ToolConsultas1.AbrirDevConsulta 5, CnnPrincipal
      If Principal.ToolConsultas1.DatSt <> "" Then
        AbrirRecorset rstUniversal, strSqlVehiculos & " WHERE IdPlaca='" & Principal.ToolConsultas1.DatSt & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          Asignar rstUniversal
        Else
          MsgBox "No se encontraron vehiculos con esta placa", vbCritical
        End If
        CerrarRecorset rstUniversal
      End If
      
      
    Case 11 'Primero
      UPrimero rstVehiculos
      Asignar rstVehiculos
    Case 12 'Anterior
      UAnterior rstVehiculos
      Asignar rstVehiculos
    Case 13 'Siguiente
      USiguiente rstVehiculos
      Asignar rstVehiculos
    Case 14 'Ultimo
      UUltimo rstVehiculos
      Asignar rstVehiculos
    Case 16 'Cerrar
      CerrarRecorset rstVehiculos
      Unload Me
    Case 17 'Actualizar
      rstVehiculos.Requery
    Case 18 'Imprimir
    
    Case 19
      If Val(TxtCampos(20).Text) <> 0 Then TxtNmMarca.Text = DevResBus("SELECT IdMarca, NmMarca From Marcas where IdMarca=" & TxtCampos(20).Text, "NmMarca", CnnPrincipal)
      If Val(TxtCampos(21).Text) <> 0 Then TxtNmColor.Text = DevResBus("SELECT IdColor, NmColor From Colores where IdColor=" & TxtCampos(21).Text, "NmColor", CnnPrincipal)
      If Val(TxtCampos(22).Text) <> 0 Then CboLineas.Text = DevResBus("SELECT IdLinea, NmLinea From Lineas where IdLinea=" & TxtCampos(22).Text, "NmLinea", CnnPrincipal)
      If Val(TxtCampos(23).Text) <> 0 Then TxtNmCarroceria.Text = DevResBus("SELECT IdCarroceria, NmCarroceria From Carrocerias where IdCarroceria=" & TxtCampos(23).Text, "NmCarroceria", CnnPrincipal)
      
      If TxtCampos(13).Text <> "" Then TxtConsulta(13).Text = DevResBus("SELECT IdTercero, concat(Nombre, ' ', Apellido1, ' ', Apellido2) As NombreCompleto From Terceros where IDTercero='" & TxtCampos(13).Text & "'", "NombreCompleto", CnnPrincipal)
      If TxtCampos(15).Text <> "" Then TxtConsulta(15).Text = DevResBus("SELECT IdTercero, concat(Nombre, ' ', Apellido1, ' ', Apellido2) As NombreCompleto From Terceros where IdTercero='" & TxtCampos(17).Text & "'", "NombreCompleto", CnnPrincipal)
      If TxtCampos(16).Text <> "" Then TxtConsulta(16).Text = DevResBus("SELECT IdTercero, concat(Nombre, ' ', Apellido1, ' ', Apellido2) As NombreCompleto From Terceros where IdTercero='" & TxtCampos(18).Text & "'", "NombreCompleto", CnnPrincipal)
  End Select
End Sub
Function Validacion() As Boolean
  If TxtCampos(2) <> "" Then
    If TxtCampos(4) <> "" Then
      If TxtCampos(5) <> "" Then
        If TxtCampos(6) <> "" Then
          If TxtCampos(7) <> "" Then
            If TxtCampos(8) <> "" Then
              If TxtCampos(9) <> "" Then
                If TxtCampos(10) <> "" Then
                  If TxtCampos(12) <> "" Then
                    If Val(TxtCampos(18).Text) <> 0 Then
                      If Val(TxtCampos(19).Text) <> 0 Then
                        If Val(TxtCampos(20).Text) <> 0 Then
                          If Val(TxtCampos(21).Text) <> 0 Then
                            If TxtCampos(13) <> "" Then
                              If TxtCampos(14) <> "" Then
                                  If TxtCampos(15) <> "" Then
                                    If TxtCampos(16) <> "" Then
                                      Validacion = True
                                    Else
                                      MsgTit "El vehiculo debe tener un dueño o propietario": Validacion = False: TxtCampos(16).SetFocus
                                    End If
                                  Else
                                    MsgTit "Debe ingresar el tenedor o persona encargada del vehiculo": Validacion = False: TxtCampos(15).SetFocus
                                  End If
                              Else
                                MsgTit "Debe especificar la poliza o el numero de soat del vehiculo": Validacion = False: TxtCampos(14).SetFocus
                              End If
                            Else
                              MsgTit "El vehiculo debe tener una empresa aseguradora": Validacion = False: TxtCampos(13).SetFocus
                            End If
                          Else
                            MsgTit "El vehiculo debe tener una carroceria": Validacion = False: TxtCampos(21).SetFocus
                          End If
                        Else
                          MsgTit "El vehiculo debe tener una linea": Validacion = False: TxtCampos(20).SetFocus
                        End If
                      Else
                        MsgTit "El vehiculo debe tener un color": Validacion = False: TxtCampos(19).SetFocus
                      End If
                    Else
                      MsgTit "El vehiculo debe tener una marca": Validacion = False: TxtCampos(18).SetFocus
                    End If
                  Else
                    MsgTit "El vehiculo debe tener un registro nacional de carga": Validacion = False: TxtCampos(12).SetFocus
                  End If
                Else
                  MsgTit "Debe ingresar la capacidad en volumen del vehiculo": Validacion = False: TxtCampos(10).SetFocus
                End If
              Else
                MsgTit "Debe ingresar la capacidad en kilos del vehiculo": Validacion = False: TxtCampos(9).SetFocus
              End If
            Else
              MsgTit "Debe especificar el peso vacío del vehiculo": Validacion = False: TxtCampos(8).SetFocus
            End If
          Else
            MsgTit "El vehiculo debe tener un numero de serie": Validacion = False: TxtCampos(7).SetFocus
          End If
        Else
          MsgTit "El vehiculo debe tener un numero de chasis": Validacion = False: TxtCampos(6).SetFocus
        End If
      Else
        MsgTit "Debe especificar el numero de ejes del vehiculo": Validacion = False: TxtCampos(5).SetFocus
      End If
    Else
      MsgTit "El vehiculo debe tener un numero de motor": Validacion = False: TxtCampos(4).SetFocus
    End If
  Else
    MsgTit "El vehiculo debe tener un modelo": Validacion = False: TxtCampos(2).SetFocus
  End If
End Function



Private Sub Timer1_Timer()
  Asignar rstVehiculos
  Timer1.Enabled = False
End Sub

Private Sub ToolVehiculos_ButtonClick(ByVal Button As MSComctlLib.Button)
    AccionTool Button.Index
End Sub



Private Sub TxtCampos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Select Case Index
      Case 13, 15, 16
        Principal.ToolConsultas1.AbrirDevConsulta 7, CnnPrincipal
        TxtCampos(Index) = Principal.ToolConsultas1.DatSt

      Case 18
        Principal.ToolConsultas1.AbrirConsultaGral "IdMarca", "NmMarca", "Marcas", CnnPrincipal
        TxtCampos(18).Text = Principal.ToolConsultas1.DatLo
        AbrirRecorset rstUniversal, "Select IdMarca, CodMinTrans FROM marcas WHERE IdMarca = " & Val(TxtCampos(18).Text), CnnPrincipal, adOpenDynamic, adLockOptimistic
        If rstUniversal.RecordCount > 0 Then
          Dim rstLineas As New ADODB.Recordset
          rstLineas.CursorLocation = adUseClient
          AbrirRecorset rstLineas, "Select IdLinea, NmLinea FROM lineas WHERE IdMarca = " & Val(TxtCampos(18).Text) & " order by NmLinea", CnnPrincipal, adOpenDynamic, adLockOptimistic
          CboLineas.ListField = "NmLinea"
          Set CboLineas.RowSource = rstLineas
        End If

      
      Case 19
        Principal.ToolConsultas1.AbrirConsultaGral "IdColor", "NmColor", "Colores", CnnPrincipal
        TxtCampos(19).Text = Principal.ToolConsultas1.DatLo
      
      Case 20
        Principal.ToolConsultas1.AbrirConsultaGral "IdLinea", "NmLinea", "Lineas", CnnPrincipal
        TxtCampos(20).Text = Principal.ToolConsultas1.DatLo
      
      Case 21
        Principal.ToolConsultas1.AbrirConsultaGral "IdCarroceria", "NmCarroceria", "Carrocerias", CnnPrincipal
        TxtCampos(21).Text = Principal.ToolConsultas1.DatLo
    End Select
  End If
End Sub

Private Sub TxtCampos_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
    Case 2, 3, 5, 8, 9, 10, 11, 13, 15, 16, 18, 19, 21
      ValidarEntrada TxtCampos(Index), KeyAscii, 1
  End Select
  If KeyAscii = 13 Then SendKeys vbTab
  
End Sub

Private Sub TxtCampos_Validate(Index As Integer, Cancel As Boolean)
  Select Case Index
    Case 13, 15, 16
      If TxtCampos(Index).Text <> "" Then
      AbrirRecorset rstUniversal, "SELECT IDTercero, Nombre, Apellido1, Apellido2 From Terceros Where IDTercero ='" & TxtCampos(Index).Text & "'", CnnPrincipal, adOpenForwardOnly, adLockReadOnly
        If rstUniversal.EOF = False Then
          TxtConsulta(Index) = rstUniversal!Nombre & " " & rstUniversal!Apellido1 & " " & rstUniversal!Apellido2 & ""
        End If
      CerrarRecorset rstUniversal
      End If
      
    Case 18
      If Val(TxtCampos(18).Text) <> 0 Then TxtNmMarca.Text = DevResBus("SELECT IdMarca, NmMarca From Marcas where IdMarca=" & TxtCampos(18).Text, "NmMarca", CnnPrincipal)
    Case 19
      If Val(TxtCampos(19).Text) <> 0 Then TxtNmColor.Text = DevResBus("SELECT IdColor, NmColor From Colores where IdColor=" & TxtCampos(19).Text, "NmColor", CnnPrincipal)
    Case 20
      If Val(TxtCampos(20).Text) <> 0 Then CboLineas.Text = DevResBus("SELECT IdLinea, NmLinea From Lineas where IdLinea=" & TxtCampos(20).Text, "NmLinea", CnnPrincipal)
    Case 21
      If Val(TxtCampos(21).Text) <> 0 Then TxtNmCarroceria.Text = DevResBus("SELECT IdCarroceria, NmCarroceria From Carrocerias where IdCarroceria=" & TxtCampos(21).Text, "NmCarroceria", CnnPrincipal)
  End Select
End Sub
Private Sub TxtFhIngreso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub
