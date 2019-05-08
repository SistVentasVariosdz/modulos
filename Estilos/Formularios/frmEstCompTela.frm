VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmEstCompTela 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estilos Componentes por Tela"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5145
      Left            =   45
      TabIndex        =   1
      Top             =   1065
      Width           =   6720
      Begin GridEX20.GridEX gexEstilosComp 
         Height          =   4785
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   8440
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         ContScroll      =   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ImageCount      =   1
         ImagePicture1   =   "frmEstCompTela.frx":0000
         DataMode        =   1
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmEstCompTela.frx":0352
         Column(2)       =   "frmEstCompTela.frx":041A
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmEstCompTela.frx":04BE
         FormatStyle(2)  =   "frmEstCompTela.frx":05F6
         FormatStyle(3)  =   "frmEstCompTela.frx":06A6
         FormatStyle(4)  =   "frmEstCompTela.frx":075A
         FormatStyle(5)  =   "frmEstCompTela.frx":0832
         FormatStyle(6)  =   "frmEstCompTela.frx":08EA
         FormatStyle(7)  =   "frmEstCompTela.frx":09CA
         FormatStyle(8)  =   "frmEstCompTela.frx":0E82
         ImageCount      =   1
         ImagePicture(1) =   "frmEstCompTela.frx":12CE
         PrinterProperties=   "frmEstCompTela.frx":1620
      End
   End
   Begin FunctionsButtons.FunctButt FunctTemporada 
      Height          =   480
      Left            =   2145
      TabIndex        =   6
      Top             =   6330
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   847
      Custom          =   $"frmEstCompTela.frx":17F8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   450
      ControlSeparator=   110
   End
   Begin VB.Label lblDes_Tela 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2685
      TabIndex        =   5
      Top             =   615
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre de Tela:"
      Height          =   225
      Left            =   870
      TabIndex        =   4
      Top             =   645
      Width           =   1755
   End
   Begin VB.Label lblCod_Tela 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2685
      TabIndex        =   3
      Top             =   240
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Código de Tela:"
      Height          =   225
      Left            =   870
      TabIndex        =   0
      Top             =   270
      Width           =   1755
   End
End
Attribute VB_Name = "frmEstCompTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SM_ESTILOS_COMPONENTES_POR_TELA()
Dim rstEstComTela As ADODB.Recordset
Dim strSQL As String
    
    Set rstEstComTela = New ADODB.Recordset
    rstEstComTela.ActiveConnection = cCONNECT
    rstEstComTela.CursorType = adOpenStatic
    rstEstComTela.CursorLocation = adUseClient
    rstEstComTela.LockType = adLockReadOnly
    strSQL = "EXEC SM_ESTILOS_COMPONENTES_POR_TELA '" & lblCod_Tela & "'"
    rstEstComTela.Open strSQL
    Set gexEstilosComp.ADORecordset = rstEstComTela
    
    With gexEstilosComp
        .Columns("ESTILO").Width = 800
        .Columns("VERSION").Width = 800
        .Columns("CONS_UNI_KILOS").Width = 1300
        .Columns("COD_COMPEST").Width = 800
        .Columns("COMPONENTE").Width = 2000
        
        .Columns("ESTILO").Caption = "Estilo"
        .Columns("VERSION").Caption = "Versión"
        .Columns("CONS_UNI_KILOS").Caption = "Cons_Uni_Kilos"
        .Columns("COD_COMPEST").Caption = "Cod_CompEst"
        .Columns("COMPONENTE").Caption = "Componente"
    End With
End Sub

Private Sub Form_Load()
'FunctTemporada.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Private Sub FunctTemporada_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo errx
Dim sSQl As String
Dim mRs As ADODB.Recordset

Select Case ActionName
    Case "ACTMASIVA"
        sSQl = "exec SG_ACTUALIZACION_MASIVA_CONSUMOS_BRUTOS  '$', '$', '$'"
        sSQl = VBsprintf(sSQl, lblCod_Tela.Caption, ComputerName(), vusu)
        Set mRs = GetRecordset(cCONNECT, sSQl)
        
        If Not mRs.EOF Then
            If mRs(0).Value = "0" Then
                MsgBox "Existen pedidos explosionados", vbCritical
                Load FrmMuestraDataActualizacionMasiva
                Set FrmMuestraDataActualizacionMasiva.DGridLista.DataSource = mRs
                FrmMuestraDataActualizacionMasiva.formato_grid
                FrmMuestraDataActualizacionMasiva.Show 1
                
                Set FrmMuestraDataActualizacionMasiva = Nothing
                
            Else
                mensaje kMESSAGE_INF_PROCESS_SATISFACTO
            End If
            Unload Me
        Else
            MsgBox "Error producido, avisar a Sistemas", vbCritical
        End If
        
    Case "SALIR"
        Unload Me
End Select

Exit Sub

errx:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
