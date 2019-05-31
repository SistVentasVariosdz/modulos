VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmModificar_Cantidades 
   Caption         =   "Modificar Cantidades"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2160
      TabIndex        =   1
      Top             =   4920
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmModificar_Cantidades.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7435
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "FrmModificar_Cantidades.frx":0097
      Column(2)       =   "FrmModificar_Cantidades.frx":015F
      FormatStylesCount=   6
      FormatStyle(1)  =   "FrmModificar_Cantidades.frx":0203
      FormatStyle(2)  =   "FrmModificar_Cantidades.frx":033B
      FormatStyle(3)  =   "FrmModificar_Cantidades.frx":03EB
      FormatStyle(4)  =   "FrmModificar_Cantidades.frx":049F
      FormatStyle(5)  =   "FrmModificar_Cantidades.frx":0577
      FormatStyle(6)  =   "FrmModificar_Cantidades.frx":062F
      ImageCount      =   0
      PrinterProperties=   "FrmModificar_Cantidades.frx":070F
   End
   Begin VB.Label Label1 
      Caption         =   "Para Modificar darle Doble click al registro Seleccionado."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "FrmModificar_Cantidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public sCod_Cliente   As String

Public sCod_PurOrd    As String

Public sCod_LotPurOrd As String

Public sCod_EstCli    As String

Public Sub Cargar()

    Dim strSql As String

    Dim y      As Integer
    
    strSql = "TG_PURORD_SM_Cantidades_Color_Talla '" & vusu & "','" & sCod_Cliente & "','" & sCod_PurOrd & "','" & sCod_LotPurOrd & "','" & sCod_EstCli & "'"
    
    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSql, cCONNECT)
    
    GridEX1.Columns(1).Caption = "Color"
    
    For y = 2 To GridEX1.Columns.count
        GridEX1.Columns(y).Width = 1000

        If y <> GridEX1.Columns.count Then
            GridEX1.Columns(y).Caption = Mid(GridEX1.Columns(y).Caption, 5, 10)
        End If

    Next
    
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, _
                                   ByVal ActionType As Integer, _
                                   ByVal ActionName As String)

    Select Case ActionName

        Case "PONERCERO"
            ''''''            If GridEX1.RowCount = 0 Then Exit Sub
            ''''''            MODIFICAR
            PONER_CERO
            
        Case "SALIR"
            Unload Me
    End Select

End Sub

Private Sub PONER_CERO()

    On Error GoTo DeleteErr

    Dim strSql     As String

    Dim ireg       As Integer

    Dim COD_COLCLI As String

    Dim COD_TALLA  As String

    Dim CANT       As String

    COD_COLCLI = Trim(GridEX1.value(GridEX1.Columns("COD_COLCLI").Index))

    ''CANT = Gridex1.value(Gridex1.col)

    If MsgBox("Desa poner en Cero todas las Cantidades de Las tallas del registro Seleccionado?", vbYesNo, "Mensaje") = vbYes Then
        ireg = GridEX1.Row

        For y = 2 To GridEX1.Columns.count - 1
            ''''Gridex1.value(Gridex1.Columns(y).Index) = 0
            COD_TALLA = Trim(GridEX1.Columns(y).Caption)
            strSql = "TG_PURORD_MOD_CANT_REQUERIDA_LOTCOLTAL '" & sCod_Cliente & "','" & sCod_PurOrd & "','" & sCod_LotPurOrd & "','" & sCod_EstCli & "','" & COD_COLCLI & "','" & COD_TALLA & "','0'"
            ExecuteCommandSQL cCONNECT, strSql
        Next

        Cargar
    
        GridEX1.Row = ireg
        MsgBox "Se Modificó Correctamante"
    End If

    Exit Sub

DeleteErr:
    errores Err.Number
End Sub

Private Sub MODIFICAR()

    On Error GoTo DeleteErr

    Dim strSql     As String

    Dim ireg       As Integer

    Dim COD_COLCLI As String

    Dim COD_TALLA  As String

    Dim CANT       As String

    For ireg = 1 To GridEX1.RowCount
        GridEX1.Row = ireg
           
    Next

    COD_COLCLI = GridEX1.value(GridEX1.Columns("COD_COLCLI").Index)
    COD_TALLA = Trim(Mid(GridEX1.Columns(GridEX1.col).Caption, 5, 10))
    CANT = GridEX1.value(GridEX1.col)
            
    strSql = "TG_PURORD_MOD_CANT_REQUERIDA_LOTCOLTAL '" & sCod_Cliente & "','" & sCod_PurOrd & "','" & sCod_LotPurOrd & "','" & sCod_EstCli & "','" & COD_COLCLI & "','" & COD_TALLA & "','" & CANT & "'"

    ireg = GridEX1.Row
 
    ExecuteCommandSQL cCONNECT, strSql
    Cargar

    GridEX1.Row = ireg
    MsgBox "Se Modificó Correctamante"

    Exit Sub

DeleteErr:
    errores Err.Number
End Sub

Private Sub GridEX1_DblClick()

    Dim ireg As Integer

    ireg = GridEX1.Row

    Load FrmModif_Cant1

    If GridEX1.col = 0 Or GridEX1.col = 1 Or GridEX1.col = GridEX1.Columns.count Then
    Else
        Set FrmModif_Cant1.oParent = Me
        FrmModif_Cant1.sCod_Cliente = sCod_Cliente
        FrmModif_Cant1.sCod_PurOrd = sCod_PurOrd
        FrmModif_Cant1.sCod_LotPurOrd = sCod_LotPurOrd
        FrmModif_Cant1.sCod_EstCli = sCod_EstCli
        FrmModif_Cant1.scod_colcli = GridEX1.value(GridEX1.Columns("COD_COLCLI").Index)
        ''FrmModif_Cant1.sCod_Talla = Trim(Mid(GridEX1.Columns(GridEX1.col).Caption, 5, 10))
        FrmModif_Cant1.sCod_Talla = Trim(GridEX1.Columns(GridEX1.col).Caption)
        FrmModif_Cant1.TxtCantidad = GridEX1.value(GridEX1.col)
        '''FrmModif_Cant1.LblTitulo.Caption = "Modificar la Cantidad de la Talla :  " & Trim(Mid(GridEX1.Columns(GridEX1.col).Caption, 5, 10))
        FrmModif_Cant1.LblTitulo.Caption = "Modificar la Cantidad de la Talla :  " & Trim(GridEX1.Columns(GridEX1.col).Caption)
        FrmModif_Cant1.Show 1
    End If

    Set FrmModif_Cant1 = Nothing
    
    Cargar
    GridEX1.Row = ireg
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then FunctButt2.SetFocus
End Sub

Private Sub GridEX1_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)

    If GridEX1.col = 1 Or GridEX1.col = GridEX1.Columns.count Then
        GridEX1.AllowEdit = False
    Else
        GridEX1.AllowEdit = True
    End If

End Sub
