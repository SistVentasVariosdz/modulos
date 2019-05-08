VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Begin VB.Form FrmRep 
   Caption         =   "Reporte de Stocks por Almacen"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   3675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   1920
      TabIndex        =   7
      Top             =   2520
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   525
      Left            =   480
      TabIndex        =   6
      Top             =   2520
      Width           =   1245
   End
   Begin VB.ComboBox CmbAlmacen 
      Height          =   315
      Left            =   1125
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2355
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Ordenamiento"
      Height          =   1935
      Left            =   225
      TabIndex        =   2
      Top             =   540
      Width           =   3465
      Begin VB.OptionButton optResOP 
         Caption         =   "Resumen x Orden de Producción"
         Height          =   240
         Left            =   195
         TabIndex        =   12
         Top             =   1560
         Visible         =   0   'False
         Width           =   2790
      End
      Begin VB.Frame fraOpcion 
         Caption         =   "Elija Filtro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   2925
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "Aceptar"
            Height          =   480
            Left            =   675
            TabIndex        =   11
            Top             =   690
            Width           =   1635
         End
         Begin VB.OptionButton optOperativos 
            Caption         =   "Operativos"
            Height          =   195
            Left            =   1620
            TabIndex        =   10
            Top             =   330
            Width           =   1095
         End
         Begin VB.OptionButton optTodos 
            Caption         =   "Todos"
            Height          =   195
            Left            =   330
            TabIndex        =   9
            Top             =   330
            Value           =   -1  'True
            Width           =   1005
         End
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Lote / Item"
         Height          =   345
         Left            =   180
         TabIndex        =   5
         Top             =   1110
         Width           =   2025
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Item / Color"
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   720
         Width           =   1845
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Proveedor / Lote / Item"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   345
         Width           =   2100
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   3270
      Top             =   2940
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Almacen:"
      Height          =   195
      Index           =   2
      Left            =   270
      TabIndex        =   1
      Top             =   165
      Width           =   660
   End
End
Attribute VB_Name = "FrmRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Reg As New ADODB.Recordset
Dim StrSql As String
Dim varCod_TipOrdTra As String

Sub GeneraReportes()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String
    
    'Aqui averiguaremos a que cod_tipordtra pertenece el almacen
    StrSql = "SELECT cod_tipordtra FROM LG_ALMACEN AL, TX_TIPOSORDTRA OT WHERE AL.TIP_ITEM = OT.TIP_ITEM AND AL.TIP_PRESENTACION = OT.TIP_PRESENTACION AND AL.Cod_Almacen = '" & Trim(Right(Me.CmbAlmacen, 6)) & "'"
    varCod_TipOrdTra = DevuelveCampo(StrSql, cConnect)
    
    Ruta = vRuta & "\Stock-Alm.xlt"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "Reporte", Reg, Left(Me.CmbAlmacen, 30), IIf(Me.Option1, "Proveedor/Lote/Item", IIf(Me.Option2, "Item/Color", "Lote/Item")), varCod_TipOrdTra, cConnect, vemp
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

Sub ResumenXOP()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String
    
    'Aqui averiguaremos a que cod_tipordtra pertenece el almacen
    StrSql = "SELECT cod_tipordtra FROM LG_ALMACEN AL, TX_TIPOSORDTRA OT WHERE AL.TIP_ITEM = OT.TIP_ITEM AND AL.TIP_PRESENTACION = OT.TIP_PRESENTACION AND AL.Cod_Almacen = '" & Trim(Right(Me.CmbAlmacen, 6)) & "'"
    varCod_TipOrdTra = DevuelveCampo(StrSql, cConnect)
    
    Ruta = vRuta & "\StockAlmResOP.xlt"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "Reporte", Reg, Left(Me.CmbAlmacen, 30), varCod_TipOrdTra, cConnect, vemp
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub


'Private Sub CmbAlmacen_Click()
'    'If Trim(Right(Me.CmbAlmacen, 6)) = "07" Then
'        LlenaCombo CmbStatusTela, "select descripcion+space(100)+flg_status_operatividad from Tx_Status_Tela_Acabada order by 1", cConnect
'        Me.Label2.Visible = True
'        Me.CmbStatusTela.Visible = True
'        optResOP.Visible = True
'    'Else
'    '    If optResOP Then Option1 = True
'    '    optResOP.Visible = False
'    '    Me.Label2.Visible = False
'    '    Me.CmbStatusTela.Visible = False
'    'End If
'End Sub

Private Sub cmdAceptar_Click()
Dim Flg_Status_Tela As String

    fraOpcion.Visible = False
    CmbAlmacen.Enabled = True
    Set Reg = Nothing
    Reg.CursorLocation = adUseClient

    'If Trim(Right(Me.CmbAlmacen, 6)) = "07" Then
'        If Trim(Right(Me.CmbStatusTela, 4)) = "" Then
'            MsgBox "Seleccion el Status de la Tela", vbInformation, Me.Caption
'            Me.CmbStatusTela.SetFocus
'            Exit Sub
'        End If
        'Flg_Status_Tela = Trim(Right(Me.CmbStatusTela, 4))
         Flg_Status_Tela = ""
    'Else
    '    Flg_Status_Tela = ""
    'End If

    VB.Screen.MousePointer = 11
    If Me.optOperativos.Value = True Then
        Reg.Open "UP_RepStocksAlmacen '" & Trim(Right(Me.CmbAlmacen, 6)) & "','" & IIf(Me.Option1, "1", IIf(Me.Option2, "2", "3")) & "','O','" & Flg_Status_Tela & "'", cConnect
    Else
        Reg.Open "UP_RepStocksAlmacen '" & Trim(Right(Me.CmbAlmacen, 6)) & "','" & IIf(Me.Option1, "1", IIf(Me.Option2, "2", "3")) & "','T','" & Flg_Status_Tela & "'", cConnect
    End If
    GeneraReportes
    VB.Screen.MousePointer = 0
    
    Set Reg = Nothing

End Sub

Private Sub Command1_Click()
Dim Flg_Status_Tela As String
    'Aqui averiguaremos a que cod_tipordtra pertenece el almacen
    StrSql = "SELECT cod_tipordtra FROM LG_ALMACEN AL, TX_TIPOSORDTRA OT WHERE AL.TIP_ITEM = OT.TIP_ITEM AND AL.TIP_PRESENTACION = OT.TIP_PRESENTACION AND AL.Cod_Almacen = '" & Trim(Right(Me.CmbAlmacen, 6)) & "'"
    varCod_TipOrdTra = DevuelveCampo(StrSql, cConnect)
    
    If varCod_TipOrdTra = "TI" And Not optResOP Then
        CmbAlmacen.Enabled = False
        fraOpcion.Visible = True
    Else
        Set Reg = Nothing
        Reg.CursorLocation = adUseClient
        
'        If Trim(Right(Me.CmbAlmacen, 6)) = "07" And Not optResOP Then
'            If Trim(Right(Me.CmbStatusTela, 4)) = "" Then
'                MsgBox "Seleccion el Status de la Tela", vbInformation, Me.Caption
'                Me.CmbStatusTela.SetFocus
'                Exit Sub
'            End If
'            Flg_Status_Tela = Trim(Right(Me.CmbStatusTela, 4))
'        Else
            Flg_Status_Tela = ""
'        End If
'
        VB.Screen.MousePointer = 11
        If optResOP Then
            Reg.Open "EXEC SM_MUESTRA_RESUMEN_X_ORDEN_PRODUCCION '" & Trim(Right(Me.CmbAlmacen, 6)) & "'", cConnect
            ResumenXOP
        Else
            Reg.Open "UP_RepStocksAlmacen '" & Trim(Right(Me.CmbAlmacen, 6)) & "','" & IIf(Me.Option1, "1", IIf(Me.Option2, "2", "3")) & "','T','" & Flg_Status_Tela & "'", cConnect
            GeneraReportes
        End If
        VB.Screen.MousePointer = 0
        
        Set Reg = Nothing
    End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Form_Load()
LlenaCombo CmbAlmacen, "Select Nom_Almacen+space(100)+ Cod_Almacen from lg_almacen  where tip_item='H' or tip_item='T' order by 1", cConnect
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

'Private Sub Option1_Click()
'    CmbStatusTela.Enabled = (Not optResOP)
'End Sub

'Private Sub Option2_Click()
'    CmbStatusTela.Enabled = (Not optResOP)
'End Sub

'Private Sub Option3_Click()
'    CmbStatusTela.Enabled = (Not optResOP)
'End Sub
'
'Private Sub optResOP_Click()
'    CmbStatusTela.Enabled = (Not optResOP)
'End Sub
