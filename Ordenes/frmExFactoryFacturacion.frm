VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmExFactoryFacturacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ex Factory Ajustada Facturación"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   975
      TabIndex        =   4
      Tag             =   "&Accept"
      Top             =   840
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2145
      TabIndex        =   3
      Tag             =   "&Cancel"
      Top             =   840
      Width           =   1170
   End
   Begin VB.Frame fraActualiza 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      TabIndex        =   1
      Top             =   -90
      Width           =   4290
      Begin MSComCtl2.DTPicker dtpFec 
         Height          =   315
         Left            =   2250
         TabIndex        =   0
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   40001
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha ExFactory Ajustada Facturacion"
         Height          =   435
         Left            =   360
         TabIndex        =   2
         Tag             =   "New PO :"
         Top             =   240
         Width           =   1770
      End
   End
End
Attribute VB_Name = "frmExFactoryFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public COD_CLIENTE As String
Public PO As String
Public LOTE_PO As String
Public ESTILO_CLIENTE As String

Private strSQL As String

Private Sub cmdAceptar_Click()
    On Error GoTo SALTO_ERROR
    Dim sFecha As String
    
    sFecha = FormatDateTime(dtpFec.value, vbShortDate)
    strSQL = "EXEC TG_LOTEST_ACTUALIZA_Fec_ExFactory_Ajustada_Facturacion '" & COD_CLIENTE & "', '" & PO & "', '" & LOTE_PO & "', '" & ESTILO_CLIENTE & "', '" & sFecha & "', '" & vusu & "'"
    Call ExecuteCommandSQL(cCONNECT, strSQL)
    MsgBox "Transacción ejectuda satisfactoriamente......", vbInformation, Me.Caption
    Unload Me
    Exit Sub
SALTO_ERROR:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

