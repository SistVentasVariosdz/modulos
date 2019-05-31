VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmShowTG_PurOrd 
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   9780
   Begin VB.TextBox txtCod_Cliente 
      Height          =   285
      Left            =   1350
      TabIndex        =   10
      Top             =   255
      Width           =   690
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2115
      TabIndex        =   9
      Top             =   255
      Width           =   2400
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   7065
      TabIndex        =   8
      Top             =   240
      Width           =   2355
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6420
      TabIndex        =   7
      Top             =   630
      Width           =   3000
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6420
      TabIndex        =   6
      Top             =   240
      Width           =   600
   End
   Begin VB.TextBox txtCod_EstCli 
      Height          =   285
      Left            =   6420
      TabIndex        =   5
      Top             =   960
      Width           =   3000
   End
   Begin VB.OptionButton optCod_PurOrd 
      Caption         =   "Purchase Order"
      Height          =   270
      Left            =   4845
      TabIndex        =   4
      Top             =   540
      Width           =   1470
   End
   Begin VB.OptionButton optCod_TemCli 
      Caption         =   "Temporada"
      Height          =   270
      Left            =   4845
      TabIndex        =   3
      Top             =   210
      Width           =   1470
   End
   Begin VB.OptionButton optCod_EstCli 
      Caption         =   "Estilo del Cliente"
      Height          =   270
      Left            =   4845
      TabIndex        =   2
      Top             =   960
      Value           =   -1  'True
      Width           =   1470
   End
   Begin SSDataWidgets_B.SSDBGrid ssgrdDatos 
      Height          =   1815
      Left            =   165
      TabIndex        =   0
      Top             =   1935
      Width           =   7470
      _Version        =   196617
      DataMode        =   2
      HeadLines       =   2
      Col.Count       =   0
      BackColorOdd    =   10354687
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      _ExtentX        =   13176
      _ExtentY        =   3201
      _StockProps     =   79
      Caption         =   "Purchase Orders"
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
   Begin MSForms.Frame fraEleccion 
      Height          =   1380
      Left            =   30
      OleObjectBlob   =   "frmShowTG_PurOrd.frx":0000
      TabIndex        =   1
      Top             =   105
      Width           =   9720
   End
End
Attribute VB_Name = "frmShowTG_PurOrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declaración de Variables Nivel Formulario
Option Explicit
Public oParent         As Object
Public sCaptionForm    As String
Public PrinterHeight
Public iLin            As Integer
Public iMante          As Integer
Dim sFlag As String

Private Sub acbOperaciones2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
End Sub

Private Sub Form_Load()
Dim x As Variant
   InitMessages
    'Call FormSet(Me, oColeccion)
    Me.Caption = sCaptionForm
    SSDBGridSetGrid0 Me.ssgrdDatos
End Sub
Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub
Private Sub acbOperaciones_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    iMante = Index
End Sub
Private Sub cmdBuscar_Click()
        Buscar
End Sub
Public Sub Buscar()
        
        Dim obj As Object 'clsTG_Participante
        Dim vBuff As Variant
        Dim iRow As Variant

        iRow = Me.ssgrdDatos.Bookmark
        Me.ssgrdDatos.Redraw = False
        
        SSDBGridSetGrid Me.ssgrdDatos
        
        'Set OBJ = New clsTG_Participante
        obj.Connect = cCONNECT
        
        RBSToSSDBGrid obj, vBuff, ssgrdDatos
        ssgrdDatos.ActiveRowStyleSet = "RowActive"
        ssgrdDatos.SelectTypeRow = ssSelectionTypeMultiSelectRange
        Set obj = Nothing
        Me.ssgrdDatos.Redraw = True
       
        Exit Sub
Errores:
    Me.MousePointer = vbDefault
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    Errores Err.Number

End Sub

Public Sub Errores(sCodigo As Long)
'Dim oCode As CodeMsg
'Dim oMessage As Object 'clsMensaje
'Dim aMess(4) As Variant
'Dim sMess As String
'Dim iPos As Integer
'
'    Select Case sCodigo
'        Case "9999"
'            oCode = kMSG_ERR_CODIGO_YA_REGISTRADO
'            Set oMessage = New clsMensaje
'            oMessage.Codigo = oCode
'            Call LoadMessage(aMess, oCode)
'            Call oMessage.ShowMsg(aMess)
'            'Aviso "El Código ya ha sido registrado.  ", 1
'
''        Case -2147217900, -2147211505
''            oCode = kMSG_ERR_REGISTRO_TIENE_TRANSAC_RELACIONADAS
''            Set omessage = New clsMensaje
''            omessage.Codigo = oCode
''            Call LoadMessage(amess, oCode)
''            Call omessage.ShowMsg(amess)
'
'            'Aviso "No se puede efectuar la operación debido a que el registro ha sido asignado a otras Tablas", 1
'        Case Else
'            sMess = Err.Description
'            iPos = InStr(1, sMess, "SERVER]", 1)
'            If iPos > 0 Then
'                sMess = Mid(sMess, iPos + 7)
'            End If
'            oCode = kMSG_ERR_HA_OCURRIDO_IMPREVISTO
'            Set oMessage = New clsMensaje
'            oMessage.Codigo = oCode
'            'oMessage.AddText = Chr(13) & " El mensaje de Error es : " & Err.Number
'            oMessage.AttribDescripLarga = Chr(13) & sMess ' Err.Description
'            Call LoadMessage(aMess, oCode)
'            Call oMessage.ShowMsg(aMess)
'
'            'Aviso "Ha ocurrido un imprevisto !!!  " & Chr(13) & _
'            'Chr(13) & "El mensaje de Error es : " & Err.Description & _
'            'Chr(13) & "El Nro. de Error es : " & Err.Number, 1
'    End Select
'
'Set oMessage = Nothing
End Sub



Sub Plin(ByVal Text)
If IsNull(Text) Then
       Text = ""
    End If
    Print #1, Text
    iLin = iLin + 1
End Sub


Private Sub Text4_Change()

End Sub
