VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmViewOPs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de las OPS"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10725
   Icon            =   "frmViewOPs.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmViewOPs.frx":01CA
   ScaleHeight     =   4560
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Detalle de las OPS"
   Begin SSDataWidgets_B.SSDBGrid ssgrdDatos 
      Height          =   3660
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   10590
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeadLines       =   2
      Col.Count       =   10
      BackColorOdd    =   10354687
      RowHeight       =   423
      Columns.Count   =   10
      Columns(0).Width=   2143
      Columns(0).Caption=   "Estilo Cliente"
      Columns(0).Name =   "COD_ESTCLI"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2143
      Columns(1).Caption=   "Cod. Est Propio"
      Columns(1).Name =   "COD_ESTPRO"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2831
      Columns(2).Caption=   "Estilo Propio"
      Columns(2).Name =   "DES_ESTPRO"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2408
      Columns(3).Caption=   "Cod. Fabrica"
      Columns(3).Name =   "COD_FABRICA"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2275
      Columns(4).Caption=   "Abr. Fabrica"
      Columns(4).Name =   "Abr_Fabrica"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2408
      Columns(5).Caption=   "Cod. Orden Prod"
      Columns(5).Name =   "COD_ORDPRO"
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   2090
      Columns(6).Caption=   "Cod. Version"
      Columns(6).Name =   "COD_VERSION"
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   2778
      Columns(7).Caption=   "Fecha Asignacion"
      Columns(7).Name =   "FEC_ASIGORD"
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   2540
      Columns(8).Caption=   "Version Costeo"
      Columns(8).Name =   "Version_Costeo"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Caption=   "Version Plan Vta"
      Columns(9).Name =   "COD_VERSION_PLAN_VENTAS"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      _ExtentX        =   18680
      _ExtentY        =   6456
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   510
      Left            =   3225
      TabIndex        =   1
      Top             =   3870
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   900
      Custom          =   $"frmViewOPs.frx":050C
      Orientacion     =   0
      Style           =   0
      Language        =   2
      TypeImageList   =   0
      ControlWidth    =   900
      ControlHeigth   =   480
      ControlSeparator=   20
   End
End
Attribute VB_Name = "frmViewOPs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Cliente As String
Public sCod_PurOrd As String
Public scod_LotPurOrd As String
Public sCod_EStCli As String
Public sCod_Fabrica As String
Public sCod_EstPro As String
Public sCod_OrdPro As String
Public oParent As Object
Public sFlag As String

Dim Strsql As String

Private Sub Form_Load()
    'Me.FunctButt2.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    FormSet Me
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ASIGNAR"
            sCod_Fabrica = Me.ssgrdDatos.Columns("Cod_Fabrica").Text
            sCod_EstPro = Me.ssgrdDatos.Columns("Cod_Estpro").Text
            sFlag = "COD_AYUDAOP"
            If Filtrar(sFlag, Me) Then
                AsignarLotesEstpro
                '          BuscarOps

            End If

        Case "DESASIGNAR"
            Desasignar

        Case "VERSIONCOSTEO"
            If Me.ssgrdDatos.Rows = 0 Then Exit Sub
            Load frmVersionCosteo
            frmVersionCosteo.sCod_Cliente = Me.sCod_Cliente
            frmVersionCosteo.sCod_PurOrd = Me.sCod_PurOrd
            frmVersionCosteo.scod_LotPurOrd = Me.scod_LotPurOrd
            frmVersionCosteo.sCod_EStCli = Me.sCod_EStCli
            frmVersionCosteo.sCod_EstPro = Me.ssgrdDatos.Columns("Cod_Estpro").Text

            Strsql = "select Tipo_orden from TG_control"
            frmVersionCosteo.Label4.Caption = Trim(DevuelveCampo(Strsql, cCONNECT)) & " :"
            frmVersionCosteo.lblPO.Caption = sCod_PurOrd
            frmVersionCosteo.lblOP.Caption = ssgrdDatos.Columns("Cod_OrdPro").Text
            frmVersionCosteo.lblEstCli.Caption = Me.ssgrdDatos.Columns("Cod_Estcli").Text
            frmVersionCosteo.lblEstPro.Caption = Me.ssgrdDatos.Columns("Cod_Estpro").Text
            frmVersionCosteo.Show 1
            BuscarOps
        Case "AVIOSADICIONALES"
            If Me.ssgrdDatos.Rows = 0 Then Exit Sub
            Load frmAviosAdicionales
            frmAviosAdicionales.sCod_Cliente = Me.sCod_Cliente
            frmAviosAdicionales.sCod_PurOrd = Me.sCod_PurOrd
            frmAviosAdicionales.scod_LotPurOrd = Me.scod_LotPurOrd
            frmAviosAdicionales.sCod_EStCli = Me.sCod_EStCli
            frmAviosAdicionales.sCod_EstPro = Me.ssgrdDatos.Columns("Cod_Estpro").Text

            Strsql = "select Tipo_orden from TG_control"
            frmAviosAdicionales.Label4.Caption = Trim(DevuelveCampo(Strsql, cCONNECT)) & " :"
            frmAviosAdicionales.lblPO.Caption = sCod_PurOrd
            frmAviosAdicionales.lblOP.Caption = ssgrdDatos.Columns("Cod_OrdPro").Text
            frmAviosAdicionales.lblEstCli.Caption = Me.ssgrdDatos.Columns("Cod_Estcli").Text
            frmAviosAdicionales.lblEstPro.Caption = Me.ssgrdDatos.Columns("Cod_Estpro").Text

            frmAviosAdicionales.CARGA_GRID
            frmAviosAdicionales.Show 1
        Case "VERSIONPLANVENTAS"
            If Me.ssgrdDatos.Rows = 0 Then Exit Sub
            frmCambioVersion.sCod_Cliente = Me.sCod_Cliente
            frmCambioVersion.sCod_PurOrd = Me.sCod_PurOrd
            frmCambioVersion.scod_LotPurOrd = Me.scod_LotPurOrd
            frmCambioVersion.sCod_EStCli = Me.sCod_EStCli
            frmCambioVersion.sCod_EstPro = Me.ssgrdDatos.Columns("Cod_Estpro").Text
            frmCambioVersion.sDes_EstPro = Me.ssgrdDatos.Columns("Des_Estpro").Text
            frmCambioVersion.Codigo = Me.ssgrdDatos.Columns("COD_VERSION_PLAN_VENTAS").Text
            frmCambioVersion.Show 1
            BuscarOps
        Case "SALIR"
            Unload Me

    End Select
End Sub

Public Sub BuscarOps()
    Dim obj As clsTG_PurOrd
    Dim vbuff As Variant
    Dim irow As Variant

    irow = Me.ssgrdDatos.Bookmark
    Me.ssgrdDatos.Redraw = False

    SSDBGridSetGrid Me.ssgrdDatos

    Set obj = New clsTG_PurOrd
    obj.ConexionString = cCONNECT
    vbuff = obj.ViewOPS(sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli)

    LibraryVBToSSDBGrid obj, vbuff, ssgrdDatos
    ssgrdDatos.ActiveRowStyleSet = "RowActive"
    ssgrdDatos.SelectTypeRow = ssSelectionTypeMultiSelectRange
    Set obj = Nothing
    Me.ssgrdDatos.Redraw = True

    Exit Sub
errores:
    Me.MousePointer = vbDefault
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    ErrorHandler Err, Err.Description

End Sub



Private Sub Asignar()
    On Error GoTo errores
    Dim obj As clsTG_PurOrd
    Dim sFlg_Modo As String


    Set obj = New clsTG_PurOrd
    obj.ConexionString = cCONNECT
    obj.AsignarOp ssgrdDatos.Columns("Cod_Fabrica").Text, ssgrdDatos.Columns("Cod_OrdPro").Text
    Set obj = Nothing

    Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
    BuscarOps
    Exit Sub
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If

    ErrorHandler Err, Err.Description

End Sub


Private Sub Desasignar()
    On Error GoTo errores
    Dim obj As clsTG_PurOrd
    Dim sFlg_Modo As String


    Set obj = New clsTG_PurOrd
    obj.ConexionString = cCONNECT
    obj.DesasignarOp sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli, ssgrdDatos.Columns("Cod_OrdPro").Text

    Set obj = Nothing

    Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
    BuscarOps
    oParent.BuscarEStilos
    Exit Sub
errores:
    If Err.Number <> 91 Then
        ErrorHandler Err, Err.Description
    Else
        Resume Next
    End If
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If

End Sub

Private Sub AsignarLotesEstpro()
    On Error GoTo errores
    Dim obj As clsTG_PurOrd
    Dim sFlg_Modo As String


    Set obj = New clsTG_PurOrd
    obj.ConexionString = cCONNECT
    obj.AsignarLotesEstpro sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, ssgrdDatos.Columns("Cod_EstCli").Text, ssgrdDatos.Columns("Cod_EstPro").Text, sCod_OrdPro
    Set obj = Nothing

    Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
    BuscarOps
    oParent.BuscarEStilos
    Exit Sub
errores:
    If Err.Number <> 91 Then
        ErrorHandler Err, Err.Description
    Else
        Resume Next
    End If
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
End Sub

