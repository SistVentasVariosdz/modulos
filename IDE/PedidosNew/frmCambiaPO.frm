VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmCambiaPO 
   Caption         =   "Cambiar clase PO"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   720
      TabIndex        =   4
      Top             =   1230
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "7~0~ACEPTAR~True~True~&Aceptar~0~0~4~~0~True~False~&Ok~~8~0~CANCELAR~True~True~&Cancelar~0~0~3~~0~False~True~&Cancel~"
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin SSDataWidgets_B.SSDBCombo cbo_ClasePO 
      Height          =   285
      Left            =   1950
      TabIndex        =   0
      Top             =   570
      Width           =   1545
      _Version        =   196617
      Columns(0).Width=   3200
      _ExtentX        =   2725
      _ExtentY        =   503
      _StockProps     =   93
      Text            =   "SSDBCombo1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Label Label1 
      Caption         =   "Nueva clase O.P.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   90
      TabIndex        =   3
      Top             =   615
      Width           =   1590
   End
   Begin VB.Label lbClaseOP 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   1335
      TabIndex        =   2
      Top             =   180
      Width           =   2250
   End
   Begin VB.Label Label4 
      Caption         =   "Clase O.P.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   1
      Top             =   165
      Width           =   1065
   End
End
Attribute VB_Name = "frmCambiaPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sCodClasePO As String
Dim mRs As New Recordset
Private Sub Cargar_ClasePO()
On Error GoTo errores
    Dim vbuff
    Dim sSQl As String
    sSQl = "exec TG_MUESTRA_CLASE_PO_PERMITIDAS "
    Set mRs = GetDataSet(cCONNECT, sSQl)
    vbuff = RowsDataSet()
    
    cbo_ClasePO.TagVariant = cbo_ClasePO.Cols
    cbo_ClasePO.RemoveAll
    LibraryVBToSSDBCombo obj, vbuff, cbo_ClasePO
    Set obj = Nothing
    
Exit Sub
errores:
    
    ErrorHandler Err, Err.Description
End Sub
 
Public Function RowsDataSet() As Variant
If Not mRs.EOF Then
 Call Refresh(mRs, vBuffProp)
 RowsDataSet = mRs.GetRows()
Else
 mRs.Close
 Set mRs = Nothing
 RowsDataSet = Empty
End If
End Function

