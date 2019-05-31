VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "*\A..\..\ECNVB6WINCTRL\ECNVB6WINCTRL.vbp"
Begin VB.Form frm001_ECNSQLUSPLIST_PARAM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relación de Parámetros"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ECNSQLUSPLIST_PARAM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin ECNVB6WINCTRL.ucLabel ucLabel1 
      Height          =   240
      Left            =   30
      TabIndex        =   2
      Top             =   3315
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   423
      Caption         =   "OUT : ES OUTPUT"
      Autosize        =   -1  'True
      ForeColor       =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ECNVB6WINCTRL.ucButton_02 btnAceptar 
      Height          =   375
      Left            =   8370
      TabIndex        =   1
      Top             =   3240
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
      Icon            =   "ECNSQLUSPLIST_PARAM.frx":058A
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshParam 
      Height          =   3195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   5636
      _Version        =   393216
      ForeColor       =   4210752
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
   Begin ECNVB6WINCTRL.ucLabel ucLabel2 
      Height          =   240
      Left            =   1560
      TabIndex        =   3
      Top             =   3315
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   423
      Caption         =   "RO : ES READONLY"
      Autosize        =   -1  'True
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ECNVB6WINCTRL.ucLabel ucLabel3 
      Height          =   240
      Left            =   3210
      TabIndex        =   4
      Top             =   3315
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   423
      Caption         =   "DF : TIENE DEFAULT"
      Autosize        =   -1  'True
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList imgL 
      Left            =   1080
      Top             =   3570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLUSPLIST_PARAM.frx":0B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLUSPLIST_PARAM.frx":1046
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm001_ECNSQLUSPLIST_PARAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PU_USPSQL_ID As String
Public PU_USPSQL_NM As String
Public PU_RS_PARAMS As ADODB.Recordset

Private Const Ki_PARAM_Col_Codigo As Integer = 0
Private Const Ki_PARAM_Col_Descripcion As Integer = 1
Private Const Ki_PARAM_Col_Tipo_ID  As Integer = 2
Private Const Ki_PARAM_Col_Tipo_NM As Integer = 3
Private Const Ki_PARAM_Col_UserTipo_ID As Integer = 4
Private Const Ki_PARAM_Col_UserTipo_NM As Integer = 5
Private Const Ki_PARAM_Col_Longitud As Integer = 6
Private Const Ki_PARAM_Col_Precision As Integer = 7
Private Const Ki_PARAM_Col_Scala As Integer = 8
Private Const Ki_PARAM_Col_IsOutput As Integer = 9
Private Const Ki_PARAM_Col_IsReadOnly As Integer = 10
Private Const Ki_PARAM_Col_IsXMLDocum As Integer = 11
Private Const Ki_PARAM_Col_HasDefValu As Integer = 12
Private Const Ki_PARAM_Col_DefaulValu As Integer = 13

Private Const Ki_Ico_Check As Integer = 1
Private Const Ki_Ico_UnCheck As Integer = 1

Private Sub btnAceptar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.Caption = "USP SQL " & PU_USPSQL_NM & " : Relación de parámetros"
End Sub

Public Sub PU_ObtenerParam()
    Dim oECNSQLHELP As ECNVB6LIB.ECNSQLHELP

    Set oECNSQLHELP = New ECNVB6LIB.ECNSQLHELP
    oECNSQLHELP.CADENA_CONEXION = GO_001_CONEXION_SQL
    Set PU_RS_PARAMS = PU_001_ObtenerParametros(oECNSQLHELP, PU_USPSQL_ID)
    Call PU_CargarInfoParam
End Sub

Public Sub PU_CargarInfoParam()
    Set mshParam.DataSource = PU_RS_PARAMS
    Call ConfiguraGrillaPARAM
End Sub

Private Sub ConfiguraGrillaPARAM()
    On Error Resume Next
    Dim F As Integer
    Dim C As Integer
    Dim i As Integer
    
    With mshParam
        .Cols = 14
        .ColWidth(Ki_PARAM_Col_Codigo) = 400
        .ColWidth(Ki_PARAM_Col_Descripcion) = 1600
        .ColWidth(Ki_PARAM_Col_Tipo_ID) = 0
        .ColWidth(Ki_PARAM_Col_Tipo_NM) = 1200
        .ColWidth(Ki_PARAM_Col_UserTipo_ID) = 0
        .ColWidth(Ki_PARAM_Col_UserTipo_NM) = 1300
        .ColWidth(Ki_PARAM_Col_Longitud) = 500
        .ColWidth(Ki_PARAM_Col_Precision) = 500
        .ColWidth(Ki_PARAM_Col_Scala) = 500
        .ColWidth(Ki_PARAM_Col_IsOutput) = 450
        .ColWidth(Ki_PARAM_Col_IsReadOnly) = 450
        .ColWidth(Ki_PARAM_Col_IsXMLDocum) = 450
        .ColWidth(Ki_PARAM_Col_HasDefValu) = 450
        .ColWidth(Ki_PARAM_Col_DefaulValu) = 1500
        
        .TextMatrix(0, Ki_PARAM_Col_Codigo) = "ID"
        .TextMatrix(0, Ki_PARAM_Col_Descripcion) = "PARAMETRO"
        .TextMatrix(0, Ki_PARAM_Col_Tipo_NM) = "TIPO"
        .TextMatrix(0, Ki_PARAM_Col_UserTipo_NM) = "TIPO USUARIO"
        .TextMatrix(0, Ki_PARAM_Col_Longitud) = "LON"
        .TextMatrix(0, Ki_PARAM_Col_Precision) = "PREC"
        .TextMatrix(0, Ki_PARAM_Col_Scala) = "SCA"
        .TextMatrix(0, Ki_PARAM_Col_IsOutput) = "OUT"
        .TextMatrix(0, Ki_PARAM_Col_IsReadOnly) = "RO"
        .TextMatrix(0, Ki_PARAM_Col_IsXMLDocum) = "XML"
        .TextMatrix(0, Ki_PARAM_Col_HasDefValu) = "DF"
        .TextMatrix(0, Ki_PARAM_Col_DefaulValu) = "DEFAULT"
        
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

            .Col = Ki_PARAM_Col_Codigo
            .CellAlignment = flexAlignCenterCenter
            .CellBackColor = &HF8F8F8
            .CellFontBold = True
        
            .Col = Ki_PARAM_Col_Descripcion
            .CellForeColor = vbBlue
            
            .Col = Ki_PARAM_Col_Tipo_NM
            .CellAlignment = flexAlignCenterCenter
            
            
            .Col = Ki_PARAM_Col_UserTipo_NM
            .CellAlignment = flexAlignCenterCenter
            
            .Col = Ki_PARAM_Col_Longitud
            .CellAlignment = flexAlignRightCenter
            
            .Col = Ki_PARAM_Col_Precision
            .CellAlignment = flexAlignRightCenter
            
            .Col = Ki_PARAM_Col_Scala
            .CellAlignment = flexAlignRightCenter
            
            .Col = Ki_PARAM_Col_IsOutput
            .CellPictureAlignment = flexAlignCenterCenter
            .CellForeColor = .CellBackColor
            i = 2: If .TextMatrix(F, .Col) = GO_ECNLIB00_CONST.VAL_CHECK Then i = 1
            Set .CellPicture = imgL.ListImages(i).Picture
            
            .Col = Ki_PARAM_Col_IsReadOnly
            .CellPictureAlignment = flexAlignCenterCenter
            .CellForeColor = .CellBackColor
            i = 2: If .TextMatrix(F, .Col) = GO_ECNLIB00_CONST.VAL_CHECK Then i = 1
            Set .CellPicture = imgL.ListImages(i).Picture
            
            .Col = Ki_PARAM_Col_IsXMLDocum
            .CellPictureAlignment = flexAlignCenterCenter
            .CellForeColor = .CellBackColor
            i = 2: If .TextMatrix(F, .Col) = GO_ECNLIB00_CONST.VAL_CHECK Then i = 1
            Set .CellPicture = imgL.ListImages(i).Picture
            
            .Col = Ki_PARAM_Col_HasDefValu
            .CellPictureAlignment = flexAlignCenterCenter
            .CellForeColor = .CellBackColor
            i = 2: If .TextMatrix(F, .Col) = GO_ECNLIB00_CONST.VAL_CHECK Then i = 1
            Set .CellPicture = imgL.ListImages(i).Picture
            
            .Col = Ki_PARAM_Col_DefaulValu
            If UCase(.TextMatrix(F, .Col)) = UCase(GO_ECNLIB00_CONST.VAL_NULL) Then
                .CellForeColor = GK_COLOR_LET_NULL
                .CellAlignment = flexAlignCenterCenter
            End If
            
        Next F
        .Row = 1
        
        .Refresh
    End With
End Sub

