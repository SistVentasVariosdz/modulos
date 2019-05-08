VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAvanEncajado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Avance Encajado"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   495
      Left            =   6165
      TabIndex        =   2
      Top             =   4410
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      Custom          =   "0~0~IMPRIMIR~Verdadero~Verdadero~&Imprimir~0~0~1~~0~Falso~Falso~&Imprimir~"
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   3990
      Left            =   60
      TabIndex        =   0
      Top             =   345
      Width           =   7350
      Begin GridEX20.GridEX gexLista 
         Height          =   3735
         Left            =   105
         TabIndex        =   1
         Top             =   135
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   6588
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAvanEncajado.frx":0000
         Column(2)       =   "frmAvanEncajado.frx":00C8
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmAvanEncajado.frx":016C
         FormatStyle(2)  =   "frmAvanEncajado.frx":02A4
         FormatStyle(3)  =   "frmAvanEncajado.frx":0354
         FormatStyle(4)  =   "frmAvanEncajado.frx":0408
         FormatStyle(5)  =   "frmAvanEncajado.frx":04E0
         FormatStyle(6)  =   "frmAvanEncajado.frx":0598
         FormatStyle(7)  =   "frmAvanEncajado.frx":0678
         ImageCount      =   0
         PrinterProperties=   "frmAvanEncajado.frx":0700
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Estilo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4215
      TabIndex        =   4
      Top             =   120
      Width           =   3105
   End
   Begin VB.Label Label1 
      Caption         =   "P.O.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   165
      TabIndex        =   3
      Top             =   105
      Width           =   3495
   End
End
Attribute VB_Name = "frmAvanEncajado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vCod_Cliente As String
Public vCod_PurOrd As String
Public vcod_lotpurord As String
Public vcod_estcli As String

Dim Strsql As String

Sub CARGA_GRID()
On Error GoTo hand

Strsql = "EXEC SM_CONSULTA_DESPACHOS_POR_PO '" & vCod_Cliente & "','" & vCod_PurOrd & "','" & vcod_lotpurord & "','" & vcod_estcli & "'"
                                
Set Me.gexLista.ADORecordset = CargarRecordSetDesconectado(Strsql, cCONNECT)
ConfigurarGrid

Exit Sub
hand:
ErrorHandler Err, "CARGA_GRID"
End Sub


Sub ConfigurarGrid()
Dim fmtCon As JSFmtCondition
Dim col As JSColumn

    With gexLista
        .Columns("tipo_1").Visible = False
        .Columns("tipo").Visible = False

        .Columns("cod_col").Width = 1000
        .Columns("Nombre_Color").Width = 2000
        .Columns("talla").Width = 900
        .Columns("Requeridas").Width = 900
        .Columns("Encajadas").Width = 900
        .Columns("% Avance").Width = 900
        
        .Columns("% Avance").Format = "#0.00"
    End With
    
    Set col = gexLista.Columns("tipo")
    Set fmtCon = gexLista.FmtConditions.Add(col.Index, jgexEqual, "2")
    fmtCon.FormatStyle.FontBold = True
'    fmtCon.FormatStyle.ForeColor = &H8000&
    
    Set col = gexLista.Columns("tipo_1")
    Set fmtCon = gexLista.FmtConditions.Add(col.Index, jgexEqual, "B")
    fmtCon.FormatStyle.FontBold = True
    fmtCon.FormatStyle.ForeColor = &HC00000
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Reporte
End Sub

Sub Reporte()
On Error GoTo ErrorImpresion

    If Me.gexLista.RowCount = 0 Then Exit Sub
    
    Dim oo As Object
    Strsql = "select ruta_logo from seguridad..seg_empresas where cod_Empresa='" & vemp1 & "'"
    
    Set oo = CreateObject("excel.application")
    'oo.Workbooks.Open vRuta & "\AvanceEncajado.xlt"
    oo.Workbooks.Open "C:\Archivos de programa\Gestion de Pedidos\AvanceEncajado.xlt"
    oo.Visible = True
    
    oo.Run "REPORTE", vCod_Cliente, vCod_PurOrd, vcod_lotpurord, vcod_estcli, DevuelveCampo(Strsql, cCONNECT), cCONNECT
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte" & Err.Description, vbCritical, "Impresion"
End Sub

