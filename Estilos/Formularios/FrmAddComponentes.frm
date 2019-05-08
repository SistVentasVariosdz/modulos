VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmAddComponentes 
   Caption         =   "Componentes"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   645
      Left            =   9120
      TabIndex        =   10
      Top             =   8040
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   926
      Custom          =   "0~0~SALIR~Verdadero~Verdadero~&Salir~0~0~1~~0~Falso~Falso~&Salir~"
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1250
      ControlHeigth   =   530
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10455
      Begin VB.TextBox TxtDes_EstPro 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   8
         Top             =   240
         Width           =   5280
      End
      Begin VB.TextBox TxtCod_EstPro 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1995
         TabIndex        =   7
         Top             =   240
         Width           =   930
      End
      Begin VB.TextBox txtDes_Version 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2715
         TabIndex        =   5
         Top             =   600
         Width           =   5445
      End
      Begin VB.TextBox TxtCod_Version 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1995
         MaxLength       =   2
         TabIndex        =   4
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Estilo Propio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   840
         TabIndex        =   9
         Tag             =   "Description:"
         Top             =   350
         Width           =   870
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   840
         TabIndex        =   6
         Tag             =   "Description:"
         Top             =   690
         Width           =   570
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   3150
      Left            =   4800
      TabIndex        =   2
      Top             =   3120
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   5556
      Custom          =   $"FrmAddComponentes.frx":0000
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   700
      ControlHeigth   =   700
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEXNo 
      Height          =   6360
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   11218
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "FrmAddComponentes.frx":00E7
      Column(2)       =   "FrmAddComponentes.frx":01AF
      FormatStylesCount=   8
      FormatStyle(1)  =   "FrmAddComponentes.frx":0253
      FormatStyle(2)  =   "FrmAddComponentes.frx":038B
      FormatStyle(3)  =   "FrmAddComponentes.frx":043B
      FormatStyle(4)  =   "FrmAddComponentes.frx":04EF
      FormatStyle(5)  =   "FrmAddComponentes.frx":05C7
      FormatStyle(6)  =   "FrmAddComponentes.frx":067F
      FormatStyle(7)  =   "FrmAddComponentes.frx":075F
      FormatStyle(8)  =   "FrmAddComponentes.frx":080B
      ImageCount      =   0
      PrinterProperties=   "FrmAddComponentes.frx":08BB
   End
   Begin GridEX20.GridEX GridEXSi 
      Height          =   6360
      Left            =   5640
      TabIndex        =   1
      Top             =   1440
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   11218
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "FrmAddComponentes.frx":0A93
      Column(2)       =   "FrmAddComponentes.frx":0B5B
      FormatStylesCount=   8
      FormatStyle(1)  =   "FrmAddComponentes.frx":0BFF
      FormatStyle(2)  =   "FrmAddComponentes.frx":0D37
      FormatStyle(3)  =   "FrmAddComponentes.frx":0DE7
      FormatStyle(4)  =   "FrmAddComponentes.frx":0E9B
      FormatStyle(5)  =   "FrmAddComponentes.frx":0F73
      FormatStyle(6)  =   "FrmAddComponentes.frx":102B
      FormatStyle(7)  =   "FrmAddComponentes.frx":110B
      FormatStyle(8)  =   "FrmAddComponentes.frx":11B7
      ImageCount      =   0
      PrinterProperties=   "FrmAddComponentes.frx":1267
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Asignados (Default)"
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
      Left            =   5640
      TabIndex        =   12
      Top             =   1200
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "No Asignados (Default)"
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
      Left            =   0
      TabIndex        =   11
      Top             =   1200
      Width           =   1980
   End
End
Attribute VB_Name = "FrmAddComponentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrSQL As String
Public vCod_EstPro As String, vCod_Version As String
Dim i As Long
Dim vCod_MotPrePro As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ADDTODO"
    If GridEXNo.RowCount = 0 Then Exit Sub
    For i = 1 To GridEXNo.RowCount
        GridEXNo.Row = i
        Call Actualizar_Componentes("I", GridEXNo)
        GridEXNo.MoveNext
    Next
    Call CARGA_GRID
Case "ADD"
    If GridEXNo.RowCount = 0 Then Exit Sub
    i = GridEXNo.Row
    Call Actualizar_Componentes("I", GridEXNo)
    Call CARGA_GRID
    GridEXNo.Row = i
Case "DEL"
    If GridEXSi.RowCount = 0 Then Exit Sub
    i = GridEXSi.Row
    Call Actualizar_Componentes("B", GridEXSi)
    Call CARGA_GRID
    GridEXSi.Row = i
Case "DELTODO"
    If GridEXSi.RowCount = 0 Then Exit Sub
    For i = 1 To GridEXSi.RowCount
        GridEXSi.Row = i
        Call Actualizar_Componentes("B", GridEXSi)
        GridEXSi.MoveNext
    Next
    Call CARGA_GRID
End Select
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Unload Me
End Sub


Sub CARGA_GRID()
StrSQL = "EXEC Es_Muestra_Componentes_No_Asignados_Default '" & vCod_EstPro & "','" & vCod_Version & "'"
Set GridEXNo.ADORecordset = CargarRecordSetDesconectado(StrSQL, cCONNECT)

GridEXNo.Columns("cod_TipCompest").Width = 0
GridEXNo.Columns("Tipo").Width = 900
GridEXNo.Columns("Codigo").Width = 800
GridEXNo.Columns("Descripcion").Width = 2900

StrSQL = "EXEC Es_Muestra_Componentes_Asignados_Default '" & vCod_EstPro & "','" & vCod_Version & "'"
Set GridEXSi.ADORecordset = CargarRecordSetDesconectado(StrSQL, cCONNECT)

GridEXSi.Columns("cod_TipCompest").Width = 0
GridEXSi.Columns("Tipo").Width = 900
GridEXSi.Columns("Codigo").Width = 800
GridEXSi.Columns("Descripcion").Width = 2900
End Sub

Sub Actualizar_Componentes(Accion As String, Grilla As Object)
On Error GoTo errComponentes
vCod_MotPrePro = DevuelveCampo("select cod_motprepro from tg_motprepro where flg_default='*'", cCONNECT)

StrSQL = "UP_Est_EstProComp '" & Accion & "','" & vCod_EstPro & "','" & vCod_Version & "','" & _
            Grilla.Value(Grilla.Columns("Codigo").Index) & "','','','1" & _
            "','" & vCod_MotPrePro & "','N'," & _
            0 & ",'',0,0,'',0,''"
            
Call ExecuteCommandSQL(cCONNECT, StrSQL)
Exit Sub
errComponentes:
    MsgBox Err.Description, vbCritical
End Sub
