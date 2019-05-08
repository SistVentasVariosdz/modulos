VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUpdateDatGenLotEst 
   Caption         =   "Actualización Datos Generales"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Update General Data"
   Begin VB.TextBox txtPrecioLOT 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1530
      TabIndex        =   9
      Text            =   "0"
      Top             =   1260
      Width           =   750
   End
   Begin VB.TextBox txtCod_DivPreLOT 
      Height          =   300
      Left            =   1530
      MaxLength       =   3
      TabIndex        =   8
      Top             =   1665
      Width           =   630
   End
   Begin VB.TextBox txtDes_DestinoLOT 
      Height          =   285
      Left            =   2190
      MaxLength       =   30
      TabIndex        =   4
      Top             =   60
      Width           =   4050
   End
   Begin VB.TextBox txtCod_DestinoLOT 
      Height          =   285
      Left            =   1530
      MaxLength       =   3
      TabIndex        =   0
      Top             =   60
      Width           =   615
   End
   Begin VB.TextBox txtPor_ComisionLOT 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1530
      TabIndex        =   2
      Text            =   "0"
      Top             =   855
      Width           =   750
   End
   Begin MSComCtl2.DTPicker dtpFec_DespachoActLOT 
      Height          =   315
      Left            =   1530
      TabIndex        =   1
      Top             =   450
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      _Version        =   393216
      Format          =   23527425
      CurrentDate     =   37159
   End
   Begin FunctionsButtons.FunctButt acbForm 
      Height          =   510
      Left            =   1920
      TabIndex        =   3
      Top             =   2250
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
   Begin VB.Label labels 
      Caption         =   "Precio"
      Height          =   255
      Index           =   14
      Left            =   60
      TabIndex        =   11
      Tag             =   "Price"
      Top             =   1290
      Width           =   1335
   End
   Begin VB.Label labels 
      Caption         =   "Division de Prenda"
      Height          =   255
      Index           =   19
      Left            =   60
      TabIndex        =   10
      Tag             =   "Garment Division"
      Top             =   1710
      Width           =   1335
   End
   Begin VB.Label labels 
      Caption         =   "Destino"
      Height          =   255
      Index           =   15
      Left            =   60
      TabIndex        =   7
      Tag             =   "Destination"
      Top             =   75
      Width           =   1200
   End
   Begin VB.Label labels 
      Caption         =   "Fecha Despacho"
      Height          =   255
      Index           =   17
      Left            =   60
      TabIndex        =   6
      Tag             =   "Delivery Date"
      Top             =   510
      Width           =   1320
   End
   Begin VB.Label labels 
      Caption         =   "Comisión"
      Height          =   255
      Index           =   18
      Left            =   60
      TabIndex        =   5
      Tag             =   "Commision"
      Top             =   870
      Width           =   1335
   End
End
Attribute VB_Name = "frmUpdateDatGenLotEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public sCod_PurORd As String
Public sCod_LotPurORd As String
Public sCod_Cliente As String
Public sCod_EstCli As String
Public sFlag As String
Public sCod_Destino As String
Public sCod_DestinoLOT As String

Private Sub acbForm_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "ACEPTAR"
        If Not ValidStep Then
            Exit Sub
        End If
        
        UpdateDatGen
    Case "CANCELAR"
        Unload Me
    End Select
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Call FormSet(Me)
End Sub

Private Function UpdateDatGen() As Boolean
On Error GoTo errores
    Dim vbuff
    Dim objPO As clsTG_LotColTal
        
    Set objPO = New clsTG_LotColTal
    objPO.Connect = cCONNECT
        
    objPO.UpdateDatGenPurORd sCod_Cliente, sCod_PurORd, sCod_LotPurORd, sCod_EstCli, txtCod_DestinoLOT.Text, CStr(dtpFec_DespachoActLOT.Value), CDbl(txtPor_ComisionLOT.Text), vusu, ComputerName, CDbl(txtPrecioLOT.Text), txtCod_DivPreLOT.Text
    
    oParent.Buscar
    oParent.BuscarEStilos
    
    Set objPO = Nothing
    Unload Me
Exit Function
errores:
    If Not objPO Is Nothing Then
        Set objPO = Nothing
    End If
    
    errores Err.Number

End Function



Private Function VAlidFechaDespacho(dFecha As String) As Boolean
On Error GoTo errores
    Dim vbuff
    Dim obj As clsTG_LotColTal
    Dim iRet As Integer
    
    Set obj = New clsTG_LotColTal
    obj.Connect = cCONNECT
    iRet = obj.VAlidFechaDespacho(dFecha)
    Set obj = Nothing
    
    If iRet = 0 Then
        VAlidFechaDespacho = True
    Else
        VAlidFechaDespacho = False
    End If
Exit Function
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    errores Err.Number
End Function

Private Sub txtCod_DestinoLOT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_DESTINOLOT"
        If Filtrar(sFlag, Me, txtCod_DestinoLOT, txtDes_DestinoLOT) Then
            Me.dtpFec_DespachoActLOT.SetFocus
        End If
    End If
End Sub
Public Function ValidStep() As Boolean
Dim aMess(4)
Dim amensaje As clsMensaje
Set amensaje = New clsMensaje
  
    If txtCod_DestinoLOT.Text = "" Then
        Mensaje CodeMsg.kMSG_ERR_NOTEMPTY
        If txtCod_DestinoLOT.Enabled Then
            Me.txtCod_DestinoLOT.SetFocus
        End If
        Exit Function
    End If

    If txtCod_DestinoLOT.Text <> "" Then
        If Not ValidCod_DestinoLot() Then
            Exit Function
        End If
    End If

    If dtpFec_DespachoActLOT.Value = "" Then
        Mensaje CodeMsg.kMSG_ERR_NOTEMPTY
        If dtpFec_DespachoActLOT.Enabled Then
            Me.dtpFec_DespachoActLOT.SetFocus
        End If
        Exit Function
    End If
    
    If Not VAlidFechaDespacho(CStr(dtpFec_DespachoActLOT.Value)) Then
        Mensaje CodeMsg.kMSG_ERR_INVALID_SELECC
        If dtpFec_DespachoActLOT.Enabled Then
            dtpFec_DespachoActLOT.SetFocus
        End If
        Exit Function
    End If
    
    If FixNulos(txtPrecioLOT.Text, vbDouble) = 0 Then
        Mensaje CodeMsg.kMSG_ERR_NOTEMPTY
        If txtPrecioLOT.Enabled Then
            Me.txtPrecioLOT.SetFocus
        End If
        Exit Function
    End If
        
    If RTrim(txtCod_DivPreLOT.Text) <> "" Then
        If Not VAlidDivPre(Me.txtCod_DivPreLOT.Text) Then
            If txtCod_DivPreLOT.Enabled Then
                txtCod_DivPreLOT.SetFocus
            End If
            Exit Function
        End If
    End If
    
    ValidStep = True
End Function

Private Function ValidCod_DestinoLot() As Boolean

    sFlag = "COD_DESTINO"
    If Not Filtrar(sFlag, Me, Me.txtCod_DestinoLOT, Me.txtDes_DestinoLOT, False) Then
        Mensaje CodeMsg.kMSG_ERR_NOTFOUND
        If Me.txtCod_DestinoLOT.Enabled Then
            Me.txtCod_DestinoLOT.SetFocus
        End If
        Exit Function
    End If

    ValidCod_DestinoLot = True
End Function


Private Function VAlidDivPre(sCod_DivPre As String) As Boolean
On Error GoTo errores
    Dim vbuff
    Dim obj As clsTG_LotColTal
    Dim bValid  As Boolean
    
    Set obj = New clsTG_LotColTal
    obj.Connect = cCONNECT
    bValid = obj.VAlidDivPre(sCod_DivPre)
    Set obj = Nothing
    
    If Not bValid Then
        Load frmDivPre
        Set frmDivPre.oParent = Me
        frmDivPre.sCod_DivPre = Me.txtCod_DivPreLOT.Text
        frmDivPre.txtCod_DivPre.Text = frmDivPre.sCod_DivPre
        frmDivPre.Show vbModal
        Set frmDivPre = Nothing
        VAlidDivPre = True
    Else
        VAlidDivPre = True
    End If
Exit Function
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    errores Err.Number
End Function

Private Sub txtCod_DivPreLOT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_DIVPRE"
            If Filtrar(sFlag, Me, txtCod_DivPreLOT, Nothing, True) Then
                acbForm.SetFocus
            Else
                If Not VAlidDivPre(Me.txtCod_DivPreLOT.Text) Then
                    Exit Sub
                Else
                    acbForm.SetFocus
                End If
            End If

    End If
End Sub
