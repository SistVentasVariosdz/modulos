VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmPreview 
   Caption         =   "Print Preview"
   ClientHeight    =   6435
   ClientLeft      =   2295
   ClientTop       =   2190
   ClientWidth     =   7005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   429
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   467
   Begin GridEX20.GEXPreview grPrev 
      Height          =   4755
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   8387
      BeginProperty ToolbarFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PageSetupText   =   "Page Set&up..."
      PrintText       =   "&Print..."
      CloseButtonText =   "&Close"
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()

    On Error Resume Next

    grPrev.Move 0, 0, ScaleWidth, ScaleHeight
    
End Sub

Private Sub grPrev_OnCloseClick()

    Unload Me
 
End Sub

Private Sub grPrev_OnPrintClick(ByVal UsePrintSetupDlg As GridEX20.JSRetBoolean)
    Unload Me
End Sub

