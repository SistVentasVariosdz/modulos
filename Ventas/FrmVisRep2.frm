VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmVisRep2 
   Caption         =   "VISTA PRELIMINAR"
   ClientHeight    =   7845
   ClientLeft      =   4665
   ClientTop       =   2385
   ClientWidth     =   7200
   Icon            =   "FrmVisRep2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   7200
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer crvRpt 
      Height          =   7755
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7200
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   0   'False
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FrmVisRep2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub crvRpt_PrintButtonClicked(UseDefault As Boolean)
'  Dim lvNumCop As Integer
'  UseDefault = False
'  On Error GoTo dprDepurar
'  frmConImp.Show vbModal
'  If gvParmL1 Then
'    With gvcrRpt
'
''      If gvIdxImp <> 81 Then
''
''         Select Case Printer.PaperSize
''           Case vbPRPSA4
''             .PaperSize = crPaperA4
''         End Select
''
''      End If
'
'      .PaperOrientation = IIf(Printer.Orientation = vbPRORPortrait, crPortrait, crLandscape)
'
'      .SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
'      .PrintOut False, gvParmN1
'    End With
'  End If
'  Exit Sub
'
'dprDepurar:
'  If Err.number = 5 Then
'    If gvParmN1 = 1 Then
'      gvcrRpt.PrintOut False
'    Else
'      For lvNumCop = 1 To gvParmN1
'        gvcrRpt.PrintOut False
'      Next lvNumCop
'    End If
'  Else
'    'Call gpCapturar_ErrorSistema("crvRpt_PrintButtonClicked - frmVisRep", teInesperado)
'  End If
End Sub

Private Sub Form_Load()
  Screen.MousePointer = vbHourglass
  With crvRpt
    .ReportSource = gvcrRpt
    .ViewReport
  End With
  'Me.Caption = gvTitMain
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
  With crvRpt
    .Top = 0
    .Left = 0
    .Height = ScaleHeight
    .Width = ScaleWidth
  End With
End Sub
