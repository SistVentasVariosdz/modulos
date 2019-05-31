VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "*\A..\..\ECNVB6WINCTRL\ECNVB6WINCTRL.vbp"
Begin VB.Form frm002_ECNSQLQRYDESIGN_TV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relación de Tablas y Vistas"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmECNSQLQRYDESIGN_TV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTab 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tablas"
      TabPicture(0)   =   "frmECNSQLQRYDESIGN_TV.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "mshTablas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vistas"
      TabPicture(1)   =   "frmECNSQLQRYDESIGN_TV.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "mshVistas"
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshTablas 
         Height          =   4305
         Left            =   30
         TabIndex        =   1
         Top             =   330
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   7594
         _Version        =   393216
         ForeColor       =   4210752
         FixedCols       =   0
         BackColorBkg    =   16777215
         GridColor       =   16053492
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshVistas 
         Height          =   4305
         Left            =   -74970
         TabIndex        =   2
         Top             =   330
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   7594
         _Version        =   393216
         ForeColor       =   4210752
         FixedCols       =   0
         BackColorBkg    =   16777215
         GridColor       =   16053492
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
   End
   Begin ECNVB6WINCTRL.ucButton_02 btnAceptar 
      Height          =   375
      Left            =   6450
      TabIndex        =   3
      Top             =   4710
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
      Icon            =   "frmECNSQLQRYDESIGN_TV.frx":0182
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
   Begin ECNVB6WINCTRL.ucButton_02 btnActualizar 
      Height          =   375
      Left            =   5100
      TabIndex        =   4
      Top             =   4710
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Icon            =   "frmECNSQLQRYDESIGN_TV.frx":071C
      Caption         =   "      &Actualizar"
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
      Left            =   7710
      TabIndex        =   5
      Top             =   4710
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
      Icon            =   "frmECNSQLQRYDESIGN_TV.frx":0CB6
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
      Left            =   0
      Top             =   4710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmECNSQLQRYDESIGN_TV.frx":1250
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmECNSQLQRYDESIGN_TV.frx":15A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmECNSQLQRYDESIGN_TV.frx":16A4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm002_ECNSQLQRYDESIGN_TV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Ki_TABLA_Col_Codigo As Integer = 0
Private Const Ki_TABLA_Col_Descripcion As Integer = 1
Private Const Ki_TABLA_Col_Creacion  As Integer = 2
Private Const Ki_TABLA_Col_Modificacion  As Integer = 3
Private Const Ki_TABLA_Col_CheckIco As Integer = 4
Private Const Ki_TABLA_Col_CheckVal As Integer = 5
Private Const Ki_TABLA_Col_CurSeleccion As Integer = 6

Private Const Ki_VISTA_Col_Codigo As Integer = 0
Private Const Ki_VISTA_Col_Descripcion As Integer = 1
Private Const Ki_VISTA_Col_Creacion  As Integer = 2
Private Const Ki_VISTA_Col_Modificacion  As Integer = 3
Private Const Ki_VISTA_Col_CheckIco As Integer = 4
Private Const Ki_VISTA_Col_CheckVal As Integer = 5
Private Const Ki_VISTA_Col_CurSeleccion As Integer = 6

Private Const Ki_Ico_Lapiz As Integer = 1
Private Const Ki_Ico_UnChk As Integer = 2
Private Const Ki_Ico_Check As Integer = 3

Private Const Ki_Tab_Tablas As Integer = 1
Private Const Ki_Tab_Vistas As Integer = 2

Private iFILA_ACTUAL_TABLA As Integer
Private iFILA_ACTUAL_VISTA As Integer

Private Sub Form_Load()
    Call PU_002_AperturarRSdeTablas
End Sub

Private Sub btnAceptar_Click()
    If PU_002_EliminarTablas = False Then Exit Sub

    Dim oMSH As MSHFlexGrid
    Dim F As Integer
    Dim i As Integer
    Dim iColCod As Integer
    Dim iColDes As Integer
    Dim iColChk As Integer
    Dim opcTipoTT As GE_TIPO_TABLA
    
    For i = 1 To 2
        If i = 1 Then
            Set oMSH = mshTablas
            iColCod = Ki_TABLA_Col_Codigo
            iColDes = Ki_TABLA_Col_Descripcion
            iColChk = Ki_TABLA_Col_CheckVal
            opcTipoTT = TT_TABLA
        Else
            Set oMSH = mshVistas
            iColCod = Ki_VISTA_Col_Codigo
            iColDes = Ki_VISTA_Col_Descripcion
            iColChk = Ki_VISTA_Col_CheckVal
            opcTipoTT = TT_VISTA
        End If
        
        With oMSH
            For F = 1 To .Rows - 1
                If .TextMatrix(F, iColChk) = GO_ECNLIB00_CONST.VAL_CHECK Then
                    Call PU_002_AgregarTabla(opcTipoTT, _
                                             .TextMatrix(F, iColCod), _
                                             .TextMatrix(F, iColDes))
                End If
            Next F
        End With
    Next i
    Unload Me
End Sub

Private Sub mshTablas_Click()
    On Error Resume Next
    
    With mshTablas
        If .Col = Ki_TABLA_Col_CheckIco Then
            Dim sValor As String
            Dim iColVal As Integer
          
            sValor = .TextMatrix(.Row, Ki_TABLA_Col_CheckVal)
            Select Case sValor
                Case GO_ECNLIB00_CONST.VAL_UNCHK: sValor = GO_ECNLIB00_CONST.VAL_CHECK
                Case GO_ECNLIB00_CONST.VAL_CHECK: sValor = GO_ECNLIB00_CONST.VAL_UNCHK
            End Select
            
            .TextMatrix(.Row, Ki_TABLA_Col_CheckVal) = sValor
            
            Select Case sValor
                Case GO_ECNLIB00_CONST.VAL_UNCHK
                    Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
                Case GO_ECNLIB00_CONST.VAL_CHECK
                    Set .CellPicture = imgL.ListImages(Ki_Ico_Check).Picture
            End Select
        End If
        .Refresh
    End With
    Call UbicaIcoEdit_Tabla
End Sub

Private Sub mshVistas_Click()
    On Error Resume Next
    
    With mshVistas
        If .Col = Ki_VISTA_Col_CheckIco Then
            Dim sValor As String
            Dim iColVal As Integer
          
            sValor = .TextMatrix(.Row, Ki_VISTA_Col_CheckVal)
            Select Case sValor
                Case GO_ECNLIB00_CONST.VAL_UNCHK: sValor = GO_ECNLIB00_CONST.VAL_CHECK
                Case GO_ECNLIB00_CONST.VAL_CHECK: sValor = GO_ECNLIB00_CONST.VAL_UNCHK
            End Select
            
            .TextMatrix(.Row, Ki_VISTA_Col_CheckVal) = sValor
            
            Select Case sValor
                Case GO_ECNLIB00_CONST.VAL_UNCHK
                    Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
                Case GO_ECNLIB00_CONST.VAL_CHECK
                    Set .CellPicture = imgL.ListImages(Ki_Ico_Check).Picture
            End Select
        End If
        .Refresh
    End With
    Call UbicaIcoEdit_Vista
End Sub

Private Sub mshTABLAS_KeyPress(KeyAscii As Integer)
    If mshTablas.Col = Ki_TABLA_Col_CheckIco Then
        Select Case KeyAscii
            Case 32, 13: Call mshTablas_Click
        End Select
    End If
End Sub

Private Sub mshVISTAS_KeyPress(KeyAscii As Integer)
    If mshVistas.Col = Ki_VISTA_Col_CheckIco Then
        Select Case KeyAscii
            Case 32, 13: Call mshVistas_Click
        End Select
    End If
End Sub

Private Sub mshTablas_RowColChange()
    Call UbicaIcoEdit_Tabla
End Sub

Private Sub mshVistas_RowColChange()
    Call UbicaIcoEdit_Vista
End Sub

Private Sub btnActualizar_Click()
    Call PU_CargarInfo
End Sub

Private Sub btnCancelar_Click()
    GO_002_ENU_OPC_WIN_RESULT = WD_CANCEL
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GO_002_SW_LOAD_DESIGN = False
End Sub

Public Sub PU_CargarInfo()
    Set mshTablas.DataSource = PU_002_CargarTablas()
    Call ConfiguraGrillaTablas
    Set mshVistas.DataSource = PU_002_CargarVistas()
    Call ConfiguraGrillaVistas
    
    Call VerificaTablasUtilizadas
End Sub

Private Sub VerificaTablasUtilizadas()
    If GO_002_RS_TABLAS.RecordCount = 0 Then Exit Sub
    
    Dim oMSH As MSHFlexGrid
    
    Dim F As Integer
    Dim i As Integer
    
    Dim iColCod As Integer
    Dim iColDes As Integer
    Dim iColChk As Integer
    Dim iColIco As Integer
    
    Dim opcTipoTT As GE_TIPO_TABLA
    Dim sCodFND As String
    
    For i = 1 To 2
        If i = 1 Then
            Set oMSH = mshTablas
            iColCod = Ki_TABLA_Col_Codigo
            iColDes = Ki_TABLA_Col_Descripcion
            iColChk = Ki_TABLA_Col_CheckVal
            iColIco = Ki_TABLA_Col_CheckIco
            opcTipoTT = TT_TABLA
        Else
            Set oMSH = mshVistas
            iColCod = Ki_VISTA_Col_Codigo
            iColDes = Ki_VISTA_Col_Descripcion
            iColChk = Ki_VISTA_Col_CheckVal
            iColIco = Ki_VISTA_Col_CheckIco
            opcTipoTT = TT_VISTA
        End If
        
        GO_002_RS_TABLAS.MoveFirst
        Do While GO_002_RS_TABLAS.EOF
            sCodFND = GO_002_RS_TABLAS.Fields(GO_002_Ks_TABLAS_CAMPO_CODIGO).Value
            sCodFND = Trim(sCodFND)
            With oMSH
                For F = 1 To .Rows - 1
                    If Trim(.TextMatrix(F, iColCod)) = sCodFND Then
                        .TextMatrix(F, iColChk) = GO_ECNLIB00_CONST.VAL_CHECK
                        .Row = F
                        .Col = iColIco
                        Set .CellPicture = imgL.ListImages(Ki_Ico_Check).Picture
                    End If
                Next F
            End With
            GO_002_RS_TABLAS.MoveNext
        Loop
    Next i
End Sub

Private Sub ConfiguraGrillaTablas()
    On Error Resume Next
    Dim F As Integer
    Dim C As Integer
    
    With mshTablas
        .Cols = 7
        .ColWidth(Ki_TABLA_Col_Codigo) = 1200
        .ColWidth(Ki_TABLA_Col_Descripcion) = 2300
        .ColWidth(Ki_TABLA_Col_Creacion) = 1800
        .ColWidth(Ki_TABLA_Col_Modificacion) = 1800
        .ColWidth(Ki_TABLA_Col_CheckIco) = 400
        .ColWidth(Ki_TABLA_Col_CheckVal) = 0
        .ColWidth(Ki_TABLA_Col_CurSeleccion) = 400
        
        
        .TextMatrix(0, Ki_TABLA_Col_Codigo) = "CODIGO"
        .TextMatrix(0, Ki_TABLA_Col_Descripcion) = "DESCRIPCION"
        .TextMatrix(0, Ki_TABLA_Col_Creacion) = "CREACION"
        .TextMatrix(0, Ki_TABLA_Col_Modificacion) = "MODIFICACION"
        
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

            .Col = Ki_TABLA_Col_Codigo
            '.CellAlignment = flexAlignCenterCenter
            .CellBackColor = &HF8F8F8
            .CellFontBold = True
        
            .Col = Ki_TABLA_Col_Descripcion
            .CellAlignment = flexAlignLeftCenter
            
            .Col = Ki_TABLA_Col_Creacion
            .CellAlignment = flexAlignCenterCenter
            
            .Col = Ki_TABLA_Col_Modificacion
            .CellAlignment = flexAlignCenterCenter
            
            .Col = Ki_TABLA_Col_CheckIco
            .CellPictureAlignment = flexAlignCenterCenter
            Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
            
            .TextMatrix(F, Ki_TABLA_Col_CheckVal) = GO_ECNLIB00_CONST.VAL_UNCHK
                        
            .Col = Ki_TABLA_Col_CurSeleccion
            Set .CellPicture = Nothing
            .CellPictureAlignment = flexAlignCenterCenter
        Next F
        .Row = 1
        
        .Refresh
    End With
End Sub

Private Sub ConfiguraGrillaVistas()
    On Error Resume Next
    Dim F As Integer
    Dim C As Integer
    
    With mshVistas
        .Cols = 7
        .ColWidth(Ki_VISTA_Col_Codigo) = 1200
        .ColWidth(Ki_VISTA_Col_Descripcion) = 2300
        .ColWidth(Ki_VISTA_Col_Creacion) = 1800
        .ColWidth(Ki_VISTA_Col_Modificacion) = 1800
        .ColWidth(Ki_VISTA_Col_CheckIco) = 400
        .ColWidth(Ki_VISTA_Col_CheckVal) = 0
        .ColWidth(Ki_VISTA_Col_CurSeleccion) = 400
        
        
        .TextMatrix(0, Ki_VISTA_Col_Codigo) = "CODIGO"
        .TextMatrix(0, Ki_VISTA_Col_Descripcion) = "DESCRIPCION"
        .TextMatrix(0, Ki_VISTA_Col_Creacion) = "CREACION"
        .TextMatrix(0, Ki_VISTA_Col_Modificacion) = "MODIFICACION"
        
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

            .Col = Ki_VISTA_Col_Codigo
            '.CellAlignment = flexAlignCenterCenter
            .CellBackColor = &HF8F8F8
            .CellFontBold = True
        
            .Col = Ki_VISTA_Col_Descripcion
            .CellAlignment = flexAlignLeftCenter
            
            .Col = Ki_VISTA_Col_Creacion
            .CellAlignment = flexAlignCenterCenter
            
            .Col = Ki_VISTA_Col_Modificacion
            .CellAlignment = flexAlignCenterCenter
            
            .Col = Ki_VISTA_Col_CheckIco
            .CellPictureAlignment = flexAlignCenterCenter
            Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
            
            .TextMatrix(F, Ki_VISTA_Col_CheckVal) = GO_ECNLIB00_CONST.VAL_UNCHK
                        
            .Col = Ki_VISTA_Col_CurSeleccion
            Set .CellPicture = Nothing
            .CellPictureAlignment = flexAlignCenterCenter
        Next F
        .Row = 1
        
        .Refresh
    End With
End Sub

Private Sub UbicaIcoEdit_Tabla()
    On Error Resume Next
    
    With mshTablas
        If .Row = 0 Then Exit Sub
        
        Dim iFilaAuxiliar As Integer
        Dim iColAnterior As Integer
        
        iColAnterior = .Col
        .Col = Ki_TABLA_Col_CurSeleccion
        
        iFilaAuxiliar = .Row
        If iFILA_ACTUAL_TABLA > 0 Then
            .Row = iFILA_ACTUAL_TABLA
            Set .CellPicture = Nothing
        End If
        .Row = iFilaAuxiliar
        Set .CellPicture = imgL.ListImages(Ki_Ico_Lapiz).Picture
                
        iFILA_ACTUAL_TABLA = .Row
        .Col = iColAnterior
        .Refresh
    End With
End Sub

Private Sub UbicaIcoEdit_Vista()
    On Error Resume Next
    
    With mshVistas
        If .Row = 0 Then Exit Sub
        
        Dim iFilaAuxiliar As Integer
        Dim iColAnterior As Integer
        
        iColAnterior = .Col
        .Col = Ki_TABLA_Col_CurSeleccion
        
        iFilaAuxiliar = .Row
        If iFILA_ACTUAL_VISTA > 0 Then
            .Row = iFILA_ACTUAL_VISTA
            Set .CellPicture = Nothing
        End If
        .Row = iFilaAuxiliar
        Set .CellPicture = imgL.ListImages(Ki_Ico_Lapiz).Picture
                
        iFILA_ACTUAL_VISTA = .Row
        .Col = iColAnterior
        .Refresh
    End With
End Sub



