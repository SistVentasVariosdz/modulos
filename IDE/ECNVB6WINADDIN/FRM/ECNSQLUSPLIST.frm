VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{196346A1-12A8-4652-B4FB-010B924E2704}#2.0#0"; "prjKEXPCheck.ocx"
Object = "*\A..\..\ECNVB6WINCTRL\ECNVB6WINCTRL.vbp"
Begin VB.Form frm001_ECNSQLUSPLIST 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relacion de Procedimientos Almacenados de Usuario"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ECNSQLUSPLIST.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPrefijoSQL 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   3300
      TabIndex        =   1
      Text            =   "USP"
      Top             =   5310
      Width           =   405
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSQLUSP 
      Height          =   5235
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   9234
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColor       =   14737632
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin prjKEXPCheck.KEXPCheck chkPrefijoSQL 
      Height          =   300
      Left            =   2100
      TabIndex        =   2
      Tag             =   "USP"
      Top             =   5310
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   529
      Caption         =   "Prefijo SQL"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckStyle      =   1
   End
   Begin ECNVB6WINCTRL.ucButton_02 btnAceptar 
      Height          =   375
      Left            =   5790
      TabIndex        =   3
      Top             =   5250
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
      Icon            =   "ECNSQLUSPLIST.frx":0A02
      Caption         =   "   &Aceptar"
      iNonThemeStyle  =   0
      BackColor       =   -2147483633
      Object.ToolTipText     =   ""
      ToolTipTitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      UseMaskColor    =   -1  'True
      RoundedBordersByTheme=   0   'False
   End
   Begin ECNVB6WINCTRL.ucButton_02 btnParam 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   5250
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Icon            =   "ECNSQLUSPLIST.frx":0F9C
      Caption         =   "      &Ver Parámetros"
      iNonThemeStyle  =   0
      BackColor       =   -2147483633
      Object.ToolTipText     =   ""
      ToolTipTitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      UseMaskColor    =   -1  'True
      RoundedBordersByTheme=   0   'False
   End
   Begin ECNVB6WINCTRL.ucButton_02 btnCancelar 
      Height          =   375
      Left            =   7050
      TabIndex        =   5
      Top             =   5250
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Icon            =   "ECNSQLUSPLIST.frx":1536
      Caption         =   "     &Cancelar"
      iNonThemeStyle  =   0
      BackColor       =   -2147483633
      Object.ToolTipText     =   ""
      ToolTipTitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      UseMaskColor    =   -1  'True
      RoundedBordersByTheme=   0   'False
   End
   Begin ComctlLib.ImageList imgL 
      Left            =   4230
      Top             =   5310
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLUSPLIST.frx":1AD0
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm001_ECNSQLUSPLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Ki_USPSQL_Col_Codigo As Integer = 0
Private Const Ki_USPSQL_Col_Descripcion As Integer = 1
Private Const Ki_USPSQL_Col_Creacion  As Integer = 2
Private Const Ki_USPSQL_Col_Modificacion  As Integer = 3
Private Const Ki_USPSQL_Col_OptionIco As Integer = 4
Private Const Ki_USPSQL_Col_OptionVal As Integer = 5
Private Const Ki_TABLA_Col_CurSeleccion As Integer = 6

Private Const Ki_Ico_Lapiz As Integer = 1

Private Const Kn_Color_OpcUnCheck As Double = &HC0C0C0
Private Const Kn_Color_OpcCheck As Double = vbBlack

Private iFILA_ACTUAL_USPSQL As Integer
Private iFILA_ULTCLIC_CHECK As Integer
Private oECNSQLHELP As ECNVB6LIB.ECNSQLHELP

Private Sub Form_Load()
    GO_001_ENU_OPC_WIN_RESULT = WD_NULL
    Set oECNSQLHELP = New ECNVB6LIB.ECNSQLHELP
    oECNSQLHELP.CADENA_CONEXION = GO_001_CONEXION_SQL
    Call CargaListaDeUSPSQL
End Sub

Private Sub mshSQLUSP_Click()
    On Error Resume Next
    
    With mshSQLUSP
        If .Col = Ki_USPSQL_Col_OptionIco Then
            Dim sValor As String
            Dim iFilaAUX As Integer
            
            iFilaAUX = .Row
            sValor = .TextMatrix(.Row, Ki_USPSQL_Col_OptionVal)
            Select Case sValor
                Case GO_ECNLIB00_CONST.VAL_UNCHK: sValor = GO_ECNLIB00_CONST.VAL_CHECK
                Case GO_ECNLIB00_CONST.VAL_CHECK: sValor = GO_ECNLIB00_CONST.VAL_UNCHK
            End Select
            .TextMatrix(.Row, Ki_USPSQL_Col_OptionVal) = sValor
            Select Case sValor
                Case GO_ECNLIB00_CONST.VAL_UNCHK
                    If iFILA_ULTCLIC_CHECK > 0 Then
                        .TextMatrix(.Row, .Col) = GO_ECNLIB00_CONST.CARESP_OPT_UNCHECK
                        .CellForeColor = Kn_Color_OpcUnCheck
                    End If
                Case GO_ECNLIB00_CONST.VAL_CHECK
                    If iFILA_ULTCLIC_CHECK > 0 Then
                        .Row = iFILA_ULTCLIC_CHECK
                        .TextMatrix(.Row, .Col) = GO_ECNLIB00_CONST.CARESP_OPT_UNCHECK
                        .CellForeColor = Kn_Color_OpcUnCheck
                    End If
                    .Row = iFilaAUX
                    .TextMatrix(.Row, .Col) = GO_ECNLIB00_CONST.CARESP_OPT_CHECKED
                    .CellForeColor = Kn_Color_OpcCheck
                    
                    iFILA_ULTCLIC_CHECK = .Row
            End Select
            .Row = iFilaAUX
        End If
        .Refresh
    End With
    Call UbicaIcoEdit_Param
End Sub

Private Sub mshSQLUSP_KeyPress(KeyAscii As Integer)
    If mshSQLUSP.Col = Ki_USPSQL_Col_OptionIco Then
        Select Case KeyAscii
            Case 32, 13: Call mshSQLUSP_Click
        End Select
    End If
End Sub

Private Sub mshSQLUSP_RowColChange()
    Call UbicaIcoEdit_Param
End Sub

Private Sub chkPrefijoSQL_Click()
'    txtPrefijoSQL.Enabled = CBool(chkPrefijoSQL.Value)
'    If chkPrefijoSQL.Value = Unchecked Then
'        txtPrefijoSQL.BackColor = Me.BackColor
'    Else
'        txtPrefijoSQL.BackColor = vbWhite
'    End If
    Call CargaListaDeUSPSQL
End Sub

Private Sub btnParam_Click()
    Call PU_ObtenerInfoDeUSPSQL(True)
End Sub

Private Sub btnAceptar_Click()
    If PU_ObtenerInfoDeUSPSQL = True Then
        GO_001_ENU_OPC_WIN_RESULT = WD_ACCEPT
        Unload Me
    Else
        mshSQLUSP.SetFocus
    End If
End Sub

Private Sub btnCancelar_Click()
    GO_001_ENU_OPC_WIN_RESULT = WD_CANCEL
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oECNSQLHELP = Nothing
End Sub

'************************************************************************************************************************************************************************************************************************************************
'PROCEDIMIENTOS DE USUARIOS LOCALES
'************************************************************************************************************************************************************************************************************************************************

Public Sub CargaListaDeUSPSQL()
    GO_001_SW_PREFIJO_SQL = CBool(chkPrefijoSQL.Value)
    GO_001_PREFIJO_USPSQL = txtPrefijoSQL.Text
        
    Dim xSQL As String
    
    xSQL = ""
    xSQL = xSQL & vbNewLine & "SELECT OBJECT_ID,"
    xSQL = xSQL & vbNewLine & "       NAME,"
    xSQL = xSQL & vbNewLine & "       CREATE_DATE = CONVERT(CHAR(10),CREATE_DATE,103) + SPACE(1) + CONVERT(CHAR(8),CREATE_DATE,108),"
    xSQL = xSQL & vbNewLine & "       MODIFY_DATE = CONVERT(CHAR(10),MODIFY_DATE,103) + SPACE(1) + CONVERT(CHAR(8),MODIFY_DATE,108)"
    xSQL = xSQL & vbNewLine & "FROM SYS.PROCEDURES"
    If GO_001_SW_PREFIJO_SQL = True Then
        xSQL = xSQL & vbNewLine & "WHERE NAME LIKE '" & GO_001_PREFIJO_USPSQL & "%'"
    End If
    xSQL = xSQL & vbNewLine & "ORDER BY NAME"
    
    Set mshSQLUSP.DataSource = oECNSQLHELP.RetornaRsCad(xSQL, ECNVB6LIB.SinParametros, False)
    Call ConfiguraGrillaUSPSQL
End Sub

Private Sub ConfiguraGrillaUSPSQL()
    On Error Resume Next
    Dim F As Integer
    Dim C As Integer
    
    With mshSQLUSP
        .Cols = 7
        .ColWidth(Ki_USPSQL_Col_Codigo) = 1200
        .ColWidth(Ki_USPSQL_Col_Descripcion) = 2300
        .ColWidth(Ki_USPSQL_Col_Creacion) = 1800
        .ColWidth(Ki_USPSQL_Col_Modificacion) = 1800
        .ColWidth(Ki_USPSQL_Col_OptionIco) = 400
        .ColWidth(Ki_USPSQL_Col_OptionVal) = 0
        .ColWidth(Ki_TABLA_Col_CurSeleccion) = 400
        
        
        .TextMatrix(0, Ki_USPSQL_Col_Codigo) = "CODIGO"
        .TextMatrix(0, Ki_USPSQL_Col_Descripcion) = "DESCRIPCION"
        .TextMatrix(0, Ki_USPSQL_Col_Creacion) = "CREACION"
        .TextMatrix(0, Ki_USPSQL_Col_Modificacion) = "MODIFICACION"
        
        .Row = 0
        .RowHeight(0) = 300
        For C = 0 To .Cols - 1
            .Col = C
            .CellBackColor = .BackColorFixed
            .CellAlignment = flexAlignCenterCenter
        Next C
                    
        For F = 1 To .Rows - 1
            .RowHeight(F) = 350
            .Row = F

            .Col = Ki_USPSQL_Col_Codigo
            .CellAlignment = flexAlignCenterCenter
            .CellBackColor = &HF8F8F8
            .CellFontBold = True
        
            .Col = Ki_USPSQL_Col_Descripcion
            .CellAlignment = flexAlignLeftCenter
            
            .Col = Ki_USPSQL_Col_Creacion
            .CellAlignment = flexAlignCenterCenter
            
            .Col = Ki_USPSQL_Col_Modificacion
            .CellAlignment = flexAlignCenterCenter
            
            .Col = Ki_USPSQL_Col_OptionIco
            .CellFontName = "Wingdings"
            .CellFontSize = "12"
            .TextMatrix(F, .Col) = GO_ECNLIB00_CONST.CARESP_OPT_UNCHECK
            .CellAlignment = flexAlignCenterCenter
            .CellForeColor = Kn_Color_OpcUnCheck
            
            .TextMatrix(F, Ki_USPSQL_Col_OptionVal) = GO_ECNLIB00_CONST.VAL_UNCHK
                        
            .Col = Ki_TABLA_Col_CurSeleccion
            Set .CellPicture = Nothing
            .CellPictureAlignment = flexAlignCenterCenter
        Next F
        .Row = 1
        
        .Refresh
    End With
End Sub

Private Sub UbicaIcoEdit_Param()
    On Error Resume Next
    
    With mshSQLUSP
        If .Row = 0 Then Exit Sub
        
        Dim iFilaAuxiliar As Integer
        Dim iColAnterior As Integer
        
        iColAnterior = .Col
        .Col = Ki_TABLA_Col_CurSeleccion
        
        iFilaAuxiliar = .Row
        If iFILA_ACTUAL_USPSQL > 0 Then
            .Row = iFILA_ACTUAL_USPSQL
            Set .CellPicture = Nothing
        End If
        .Row = iFilaAuxiliar
        Set .CellPicture = imgL.ListImages(Ki_Ico_Lapiz).Picture
                
        iFILA_ACTUAL_USPSQL = .Row
        .Col = iColAnterior
        .Refresh
    End With
End Sub

Public Function PU_ObtenerInfoDeUSPSQL(Optional ByVal blSW_VerInfo As Boolean = False) As Boolean
    PU_ObtenerInfoDeUSPSQL = False
    If frm001_ECNSQLUSPLIST.Visible = False Then
        MsgBox "La ventana [LISTADO DE USP SQL] debe encontrarse activa...", vbInformation, Me.Caption
        Exit Function
    End If
    If iFILA_ULTCLIC_CHECK = 0 Then
        MsgBox "No ha seleccionado ningun registro...", vbInformation, Me.Caption
        Exit Function
    End If
    
    GO_001_SW_PREFIJO_SQL = CBool(chkPrefijoSQL.Value)
    GO_001_PREFIJO_USPSQL = txtPrefijoSQL.Text
    
    With mshSQLUSP
        GO_001_USPSQL_SEL_COD = .TextMatrix(iFILA_ULTCLIC_CHECK, Ki_USPSQL_Col_Codigo)
        GO_001_USPSQL_SEL_NOM = .TextMatrix(iFILA_ULTCLIC_CHECK, Ki_USPSQL_Col_Descripcion)
    End With
   
    Set GO_001_LST_PARAMETROS = New ADODB.Recordset
    Set GO_001_LST_PARAMETROS.DataSource = PU_001_ObtenerParametros(oECNSQLHELP, GO_001_USPSQL_SEL_COD)
    PU_ObtenerInfoDeUSPSQL = True
    If blSW_VerInfo = True Then
        If GO_001_LST_PARAMETROS.RecordCount = 0 Then
            MsgBox "El SP SQL [" & GO_001_USPSQL_SEL_NOM & "] no contiene parámetros...", vbInformation, Me.Caption
            Exit Function
        End If
        Load frm001_ECNSQLUSPLIST_PARAM
        With frm001_ECNSQLUSPLIST_PARAM
            .PU_USPSQL_ID = GO_001_USPSQL_SEL_COD
            .PU_USPSQL_NM = GO_001_USPSQL_SEL_NOM
            Set .PU_RS_PARAMS = GO_001_LST_PARAMETROS
            Call .PU_CargarInfoParam
            .Show 1
        End With
        Set frm001_ECNSQLUSPLIST_PARAM = Nothing
    End If
End Function
