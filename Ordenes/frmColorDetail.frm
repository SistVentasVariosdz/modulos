VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmColorDetail 
   Caption         =   "Detalle por Color"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Colour Detail"
   Begin SSDataWidgets_B.SSDBGrid ssgrdDatos 
      Height          =   6570
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   10170
      _Version        =   196617
      DataMode        =   2
      HeadLines       =   2
      Col.Count       =   11
      BackColorOdd    =   10354687
      RowHeight       =   423
      Columns.Count   =   11
      Columns(0).Width=   2831
      Columns(0).Caption=   "Color"
      Columns(0).Name =   "Cod_ColCli"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1852
      Columns(1).Caption=   "Prendas Requeridas"
      Columns(1).Name =   "Num_PreReq"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   5
      Columns(1).FieldLen=   256
      Columns(2).Width=   1720
      Columns(2).Caption=   "Prendas Despachadas"
      Columns(2).Name =   "Num_PreDes"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   5
      Columns(2).FieldLen=   256
      Columns(3).Width=   2223
      Columns(3).Caption=   "Imp. Total Requerido"
      Columns(3).Name =   "Imp_TotalPre"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   5
      Columns(3).FieldLen=   256
      Columns(4).Width=   2461
      Columns(4).Caption=   "Imp.Total Despachado"
      Columns(4).Name =   "Imp_TotalDes"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   0
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   5
      Columns(4).FieldLen=   256
      Columns(5).Width=   1244
      Columns(5).Caption=   "Motivo Atraso"
      Columns(5).Name =   "Cod_MotAtr"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   2858
      Columns(6).Caption=   "Descrip. Motivo Atraso"
      Columns(6).Name =   "Des_MotAtr"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1085
      Columns(7).Caption=   "Ano"
      Columns(7).Name =   "Ano"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1005
      Columns(8).Caption=   "Mes"
      Columns(8).Name =   "Mes"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Caption=   "Serie Factura"
      Columns(9).Name =   "Cod_SerFac"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Caption=   "Factura"
      Columns(10).Name=   "Cod_Factura"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      _ExtentX        =   17939
      _ExtentY        =   11589
      _StockProps     =   79
      Caption         =   "Color Detail"
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
   Begin FunctionsButtons.FunctButt acbForm 
      Height          =   510
      Left            =   3885
      TabIndex        =   1
      Top             =   6645
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "7~0~ACEPTAR~True~True~&Aceptar~0~0~4~~0~True~False~&Ok~~8~0~CANCELAR~True~True~&Cancelar~0~0~3~~0~False~True~&Cancel~"
      Orientacion     =   0
      Style           =   1
      Language        =   1
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmColorDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent         As Object
Public sCaptionForm    As String
Public PrinterHeight
Public iLin            As Integer
Public iMante          As Integer
Public sCod_Cliente    As String
Public sCod_PurOrd     As String
Public scod_LotPurOrd  As String
Public sCod_EStCli     As String
Public dPor_ComisionCliente As Double

Dim sFlag As String


Private Sub acbForm_ActionClick(ByVal index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            Unload Me
        Case "CANCELAR"
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
Dim x As Variant
    'InitMessages C.A.R.
    Me.Caption = sCaptionForm
    SSDBGridSetGrid0 Me.ssgrdDatos
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub


Public Function Buscar() As Boolean
        
        Dim obj As clsTG_LotColTal
        Dim vbuff As Variant
        Dim irow As Variant

        Buscar = False
        
        irow = Me.ssgrdDatos.Bookmark
        Me.ssgrdDatos.Redraw = False
        
        SSDBGridSetGrid Me.ssgrdDatos
        
        Set obj = New clsTG_LotColTal
        obj.ConexionString = cCONNECT
        
        vbuff = obj.ViewVectorColorKey(sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli)
        
        LibraryVBToSSDBGrid obj, vbuff, ssgrdDatos
        ssgrdDatos.ActiveRowStyleSet = "RowActive"
        ssgrdDatos.SelectTypeRow = ssSelectionTypeMultiSelectRange
        
        
        
        Set obj = Nothing
        SSDBGridTOTALES ssgrdDatos
        
        Me.ssgrdDatos.Redraw = True
        If Me.Enabled And Me.Visible Then
            Me.ssgrdDatos.SetFocus
        End If
        
        Exit Function
errores:
    Me.MousePointer = vbDefault
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    ErrorHandler Err, Err.Description

End Function


