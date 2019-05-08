VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmDetalleServTenido 
   Caption         =   "Detalle de Servicio de Teñido"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   13260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   525
      Left            =   11970
      TabIndex        =   0
      Top             =   3540
      Width           =   1245
   End
   Begin GridEX20.GridEX gexLista 
      Height          =   3450
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   6085
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmDetalleServTenido.frx":0000
      Column(2)       =   "frmDetalleServTenido.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmDetalleServTenido.frx":016C
      FormatStyle(2)  =   "frmDetalleServTenido.frx":02A4
      FormatStyle(3)  =   "frmDetalleServTenido.frx":0354
      FormatStyle(4)  =   "frmDetalleServTenido.frx":0408
      FormatStyle(5)  =   "frmDetalleServTenido.frx":04E0
      FormatStyle(6)  =   "frmDetalleServTenido.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmDetalleServTenido.frx":0678
   End
End
Attribute VB_Name = "frmDetalleServTenido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sSer_OrdCom As String
Public sCod_OrdCom As String
Public sCod_Tela As String
Public sCod_Combinacion As String
Public sDes_Tela As String

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Public Sub CARGA_GRID()
Dim Strsql  As String
    'Esta cadena es para devolver el Codigo de Cliente
    Strsql = "EXEC SM_STOCKS_EN_SERVICIO_TENIDO_DETALLE '$','$','$','$'"
    Strsql = VBsprintf(Strsql, sSer_OrdCom, sCod_OrdCom, sCod_Tela, sCod_Combinacion)
    
    Set gexLista.ADORecordset = CargarRecordSetDesconectado(Strsql, cCONNECT)
    
    SetGeneralGridEX gexLista, 0, 1
    
    Call CONFIGURAR_GRID
    
    Me.Caption = Me.Caption & " O/C: " & sSer_OrdCom & "-" & sCod_OrdCom & "      Tela: " & sCod_Tela & " " & sDes_Tela & "        " & IIf(RTrim(sCod_Combinacion) = "", "", "Comb: " & sCod_Combinacion)
 
End Sub

Public Sub CONFIGURAR_GRID()
    
    Me.gexLista.Columns("NUM_MOVSTK").Caption = "Nro Movim"
    Me.gexLista.Columns("FEC_MOVSTK").Caption = "Fecha"
    Me.gexLista.Columns("NUM_GUIA").Caption = "N.Guia"
    Me.gexLista.Columns("ENVIADO").Caption = "Enviado"
    Me.gexLista.Columns("INGRESADO").Caption = "Ingresado"
    Me.gexLista.Columns("ROLLOS_ENVIADOS").Caption = "Rollos Enviados"
    Me.gexLista.Columns("ROLLOS_RECIBIDOS").Caption = "Rollos Recibidos"
    Me.gexLista.Columns("MOVIMIENTO").Caption = "Movimiento"
    Me.gexLista.Columns("OBSERVACIONES").Caption = "Observaciones"
    
    gexLista.Columns("NUM_MOVSTK").Width = 870
    gexLista.Columns("FEC_MOVSTK").Width = 1005
    gexLista.Columns("NUM_GUIA").Width = 765
    gexLista.Columns("ENVIADO").Width = 810
    gexLista.Columns("INGRESADO").Width = 915
    gexLista.Columns("ROLLOS_ENVIADOS").Width = 1260
    gexLista.Columns("ROLLOS_RECIBIDOS").Width = 1320
    gexLista.Columns("MOVIMIENTO").Width = 2500
    gexLista.Columns("OBSERVACIONES").Width = 3500
End Sub

Public Function SetGeneralGridEX(ByRef GridEx As GridEX20.GridEx, ByVal iFixsCols As Integer, ByVal iTipoColorBack As Integer)

    If iFixsCols > 0 Then
        GridEx.FrozenColumns = iFixsCols
    End If
    
    If iTipoColorBack = 1 Then
        GridEx.BackColor = &H80000018
        GridEx.BackColorBkg = &H80000018
        GridEx.GridLines = jgexGLVertical
        GridEx.GridLineStyle = jgexGLSSmallDots
    Else
        GridEx.BackColor = &H80000005
        GridEx.BackColorBkg = &H80000005
        GridEx.GridLines = jgexGLBoth
        GridEx.GridLineStyle = jgexGLSSmallDots
    End If
    
End Function



Private Sub gexLista_DblClick()
    Dim i As Integer
    For i = 1 To gexLista.Columns.Count
        Debug.Print gexLista.Name & ".Columns(" & Chr(34) & gexLista.Columns(i).Caption & Chr(34) & ").width = " & CStr(gexLista.Columns(i).Width)
    Next

End Sub
