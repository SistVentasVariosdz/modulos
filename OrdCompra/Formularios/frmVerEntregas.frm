VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVerEntregas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entregas"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   495
      Left            =   6210
      TabIndex        =   3
      Top             =   4080
      Width           =   1260
   End
   Begin VB.Frame FraLista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3990
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9075
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   3630
         Left            =   180
         TabIndex        =   2
         Top             =   225
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   6403
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   17
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   2
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   7590
      TabIndex        =   0
      Top             =   4080
      Width           =   1260
   End
End
Attribute VB_Name = "frmVerEntregas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Strsql As String
Dim Rs_Lista As ADODB.Recordset
Dim sTipo As String

'Definicion de variables que seran pasadas por nuestro master
Public varSer_OrdComp As String, varCod_OrdComp As String, varSec_OrdComp As String

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Sub CARGA_GRID()
    Dim Tipo_Consulta As String
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = cConnect
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    'Esta cadena es para devolver el Codigo de Cliente
    Strsql = "EXEC SM_ACT_ENTREGAS_OC '" & varSer_OrdComp & "','" & varCod_OrdComp & "','" & varSec_OrdComp & "'"
    
    Rs_Lista.Open Strsql
    Set DGridLista.DataSource = Rs_Lista
    DGridLista.Refresh

End Sub


Private Sub Command1_Click()
    Call REPORTE
End Sub

Private Sub Form_Load()
 Call FormateaGrid(DGridLista)
End Sub



Public Sub REPORTE()
On Error GoTo ErrorImpresion
    Dim oo As Object
    Set oo = CreateObject("excel.application")
    
    oo.Workbooks.Open vRuta & "\SeguimientoOrdComp.xlt"
    oo.Visible = True
    oo.Run "REPORTE", varSer_OrdComp, varCod_OrdComp, varSec_OrdComp, cConnect
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte de Seguimiento de O/C" & Err.Description, vbCritical, "Impresion"
End Sub


