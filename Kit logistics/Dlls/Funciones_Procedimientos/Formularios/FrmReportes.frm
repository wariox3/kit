VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmReportes 
   Caption         =   "Reportes..."
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11115
   Icon            =   "FrmReportes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWERLibCtl.CRViewer CRViewer 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FrmReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Private mflgContinuar As Boolean
Private mstrParametro1 As String
Private mlngParametro2 As Long
Public Sub PasarParametros(sParam1 As String, lParam2 As Long)
    mstrParametro1 = sParam1
    mlngParametro2 = lParam2
End Sub

Private Sub Form_Activate()
  If Not mflgContinuar Then Unload Me
End Sub

Private Sub Form_Load()
    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
    On Error GoTo ErrHandler
    'Abrir el reporte
    Screen.MousePointer = vbHourglass
    mflgContinuar = True
    Set crReport = crApp.OpenReport(RutaInf, 1)
    crReport.DiscardSavedData
    crReport.Database.SetDataSource rstFunPro
    'crReport.SQLQueryString = FuenteSQL
    
    ' Parametros del reporte
    Me.Caption = TituloVentana
    Set crParamDefs = crReport.ParameterFields
    For Each crParamDef In crParamDefs
        Select Case crParamDef.ParameterFieldName
            Case "Parametro1"
                crParamDef.AddCurrentValue (mstrParametro1)
            Case "Parametro2"
                crParamDef.AddCurrentValue (mlngParametro2)
        End Select
    Next
    If TituloInf <> "" Then crReport.ReportTitle = TituloInf
    CRViewer.ReportSource = crReport
    CRViewer.DisplayGroupTree = False
    
    Select Case OpcReporte
      Case 1
        CRViewer.PrintReport
      Case 2
        CRViewer.ViewReport
    End Select
    Screen.MousePointer = vbDefault
    Set crParamDefs = Nothing
    Set crParamDef = Nothing
    Exit Sub

ErrHandler:
    If Err.Number = -2147206461 Then
        MsgBox "El archivo de reporte no se encuentra, restáurelo de los discos de instalación", vbCritical
    Else
        MsgBox Err.Description, vbCritical + vbOKOnly
    End If
    mflgContinuar = False
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Resize()
    CRViewer.Top = 0
    CRViewer.Left = 0
    CRViewer.Height = ScaleHeight
    CRViewer.Width = ScaleWidth
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set crReport = Nothing
    Set crApp = Nothing
End Sub
