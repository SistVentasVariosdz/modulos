VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmActPrecioOC 
   Caption         =   "Actualizar Precio O/C Tejeduria"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1920
      TabIndex        =   1
      Top             =   1800
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmActPrecioOC.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame2 
      Height          =   680
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   6375
      Begin VB.TextBox TxtPrecio 
         Height          =   285
         Left            =   3240
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   4
         Top             =   330
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6375
      Begin VB.Label LblOC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Orden Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   600
         Width           =   1560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Orden Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label LblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion Cliente "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   3840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   600
      End
   End
End
Attribute VB_Name = "FrmActPrecioOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public vCod_Cliente As String, vSer_OrdComp As String, vCod_OrdComp As String, vSec_OrdComp As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACTUALIZAR"
    Call ACTUALIZAR
Case "SALIR"
    Unload Me
End Select
End Sub

Sub ACTUALIZAR()
On Error GoTo errActualizar

strSQL = "Ventas_Up_Actualiza_Precio_Tejeduria '" & vCod_Cliente & "','" & vSer_OrdComp & "','" & vCod_OrdComp & "','" & vSec_OrdComp & "'," & CDbl(TxtPrecio.Text)
ExecuteCommandSQL cCONNECT, strSQL
Unload Me
Exit Sub
errActualizar:
    ErrorHandler Err, "Actualizar Precio"
End Sub

Private Sub TxtPrecio_GotFocus()
SelectionText TxtPrecio
End Sub

Private Sub TxtPrecio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtPrecio, KeyAscii, True, 3)
End If
End Sub
