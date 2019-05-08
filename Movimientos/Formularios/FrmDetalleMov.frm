VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmDetalleMov 
   Caption         =   "Detalle de Movimiento"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   11325
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11295
      Begin VB.Label LblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   4560
         TabIndex        =   17
         Top             =   480
         Width           =   1485
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Color"
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
         Left            =   3960
         TabIndex        =   16
         Top             =   525
         Width           =   450
      End
      Begin VB.Label LblCod_EstCli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   9720
         TabIndex        =   15
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label LblDestino 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   7200
         TabIndex        =   14
         Top             =   480
         Width           =   765
      End
      Begin VB.Label LblTalla 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   3240
         TabIndex        =   13
         Top             =   480
         Width           =   645
      End
      Begin VB.Label LblComb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   480
         Width           =   1485
      End
      Begin VB.Label LblDes_Item 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   4560
         TabIndex        =   11
         Top             =   150
         Width           =   6165
      End
      Begin VB.Label LblCod_Item 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   3240
         TabIndex        =   10
         Top             =   150
         Width           =   1245
      End
      Begin VB.Label LblNP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   150
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cod. EstCli."
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
         Left            =   8640
         TabIndex        =   8
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
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
         Left            =   6360
         TabIndex        =   7
         Top             =   525
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Talla"
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
         Left            =   2640
         TabIndex        =   6
         Top             =   525
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Comb"
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
         Left            =   480
         TabIndex        =   5
         Top             =   525
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Item"
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
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NP"
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
         Left            =   480
         TabIndex        =   3
         Top             =   270
         Width           =   270
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   8760
      TabIndex        =   1
      Top             =   5040
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmDetalleMov.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX gexList 
      Height          =   4005
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   7064
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      TabKeyBehavior  =   1
      SelectionStyle  =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "FrmDetalleMov.frx":0090
      FormatStyle(2)  =   "FrmDetalleMov.frx":01C8
      FormatStyle(3)  =   "FrmDetalleMov.frx":0278
      FormatStyle(4)  =   "FrmDetalleMov.frx":032C
      FormatStyle(5)  =   "FrmDetalleMov.frx":0404
      FormatStyle(6)  =   "FrmDetalleMov.frx":04BC
      FormatStyle(7)  =   "FrmDetalleMov.frx":059C
      ImageCount      =   0
      PrinterProperties=   "FrmDetalleMov.frx":05BC
   End
End
Attribute VB_Name = "FrmDetalleMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Fabrica As String, sCod_OrdPro As String, sCOD_ITEM As String, sCOD_COMB As String, scod_color As String
Public sCOD_TALLA As String, sCod_Destino As String, sCod_EstCli As String

Dim StrSql As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    Call Reporte
Case "SALIR"
    Unload Me
End Select
End Sub

Sub CARGA_GRID()
StrSql = "lg_muestra_detalle_despachos_avios_almacen_01 '" & sCod_Fabrica & "','" & sCod_OrdPro & "','" & sCOD_ITEM & "','" & _
                                        sCOD_COMB & "','" & scod_color & "','" & sCOD_TALLA & "','" & sCod_Destino & "','" & sCod_EstCli & "'"
VB.Screen.MousePointer = 11
Set Me.gexList.ADORecordset = CargarRecordSetDesconectado(StrSql, cConnect)

gexList.Columns("tipo_movimiento").Width = 3000
gexList.Columns("cod_proveedor").Width = 0

VB.Screen.MousePointer = 0
End Sub

Public Sub Reporte()
On Error GoTo ErrorImpresion
    Dim oo As Object
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\RptDetMovxAvios.XLT"
    oo.Visible = True
    
    oo.Run "REPORTE", LblNP, LblCod_Item & "-" & Trim(LblDes_Item), Trim(LblComb), LblColor, LblTalla, LblDestino, LblCod_EstCli, StrSql, cConnect
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte  " & err.Description, vbCritical, "Impresion"
End Sub

