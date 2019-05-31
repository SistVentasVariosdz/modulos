VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmImprimirFacturasPacking 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   4200
      TabIndex        =   0
      Top             =   4680
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~SALIR~Verdadero~Verdadero~&Salir~0~0~1~~0~Falso~Falso~&Salir~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX grdImpresion 
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8070
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmImprimirFacturasPacking.frx":0000
      Column(2)       =   "frmImprimirFacturasPacking.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmImprimirFacturasPacking.frx":016C
      FormatStyle(2)  =   "frmImprimirFacturasPacking.frx":02A4
      FormatStyle(3)  =   "frmImprimirFacturasPacking.frx":0354
      FormatStyle(4)  =   "frmImprimirFacturasPacking.frx":0408
      FormatStyle(5)  =   "frmImprimirFacturasPacking.frx":04E0
      FormatStyle(6)  =   "frmImprimirFacturasPacking.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmImprimirFacturasPacking.frx":0678
   End
End
Attribute VB_Name = "frmImprimirFacturasPacking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public sCod_Cliente   As String

Public sCod_PurOrd    As String

Public sCod_LotPurOrd As String

Public sCod_EstCli    As String

Public sStore         As String

Private Sub Form_Load()

    On Error GoTo Err_Buscar

    Dim oFrm As New Frm_Toolbar

    oFrm.CambiarContenedor Me
    Set oFrm = Nothing

    Exit Sub

Err_Buscar:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Imprimir"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, _
                                   ByVal ActionType As Integer, _
                                   ByVal ActionName As String)

    Select Case ActionName

        Case "SALIR"
            Unload Me
    End Select

End Sub

Public Sub CARGAGRILLA()

    On Error GoTo Err_Buscar

    strSql = "EXEC " & sStore & " '" & sCod_Cliente & "', '" & sCod_PurOrd & "','" & sCod_LotPurOrd & "','" & sCod_EstCli & "'"

    Set grdImpresion.ADORecordset = CargarRecordSetDesconectado(strSql, cCONNECT)
    
    '    Gridex1.Columns("Fec_Trabajo").Width = 1200
    '    Gridex1.Columns("Fec_Trabajo").Caption = "Fecha"
    '
    '    Gridex1.Columns("Cod_Maquina").Width = 1000
    '    Gridex1.Columns("Cod_Maquina").Caption = "Maquina"
    '
    '    Gridex1.Columns("Sec").Width = 500
    '
    '    Gridex1.Columns("Kilos_Brutos").Width = 900
    '    Gridex1.Columns("Kilos_Brutos").Caption = "Kg Brutos"
    '
    '    Gridex1.Columns("Tara").Width = 900
    '
    '    Gridex1.Columns("Kilos_Netos").Width = 900
    '    Gridex1.Columns("Kilos_Netos").Caption = "Kg Netos"
    '
    '    Gridex1.Columns("Numero_Cajas").Width = 800
    '    Gridex1.Columns("Numero_Cajas").Caption = "Nº Cajas"
    '
    '    Gridex1.Columns("Husos_Inactivos").Width = 1000
    '    Gridex1.Columns("Husos_Inactivos").Caption = "Husos Inac"
    '
    '    Gridex1.Columns("Canilla").Width = 1000
    '    Gridex1.Columns("Canilla").Caption = "Canilla"
    
    Exit Sub

Err_Buscar:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Imprimir"
End Sub
