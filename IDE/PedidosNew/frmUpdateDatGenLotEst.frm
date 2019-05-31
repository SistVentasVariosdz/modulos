VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUpdateDatGenLotEst 
   Caption         =   "Actualización Datos Generales"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Update General Data"
   Begin VB.TextBox txtUtilidadCotizada 
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4620
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "0"
      Top             =   1920
      Width           =   750
   End
   Begin VB.TextBox txtPrecio_Cotizado 
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4620
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "0"
      Top             =   1680
      Width           =   750
   End
   Begin VB.TextBox TxtLead 
      Height          =   285
      Left            =   1560
      TabIndex        =   29
      Top             =   4920
      Width           =   4695
   End
   Begin VB.TextBox Txtcod_Lead 
      Height          =   285
      Left            =   1080
      TabIndex        =   28
      Top             =   4920
      Width           =   495
   End
   Begin VB.OptionButton optPrePackNo 
      Caption         =   "No"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3120
      TabIndex        =   26
      Top             =   4560
      Width           =   975
   End
   Begin VB.OptionButton optPrePackSi 
      Caption         =   "Si"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1680
      TabIndex        =   25
      Top             =   4560
      Width           =   1095
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
      Left            =   4620
      TabIndex        =   22
      Text            =   "0"
      Top             =   1095
      Width           =   750
   End
   Begin VB.OptionButton optComisionEnPorcentaje 
      Caption         =   "En Porcentaje"
      Height          =   240
      Left            =   1530
      TabIndex        =   19
      Top             =   1080
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton optComisionEnImporte 
      Caption         =   "En Importe"
      Height          =   240
      Left            =   1545
      TabIndex        =   18
      Top             =   1380
      Width           =   1335
   End
   Begin VB.TextBox txtImp_Comision 
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
      Left            =   4620
      TabIndex        =   17
      Text            =   "0"
      Top             =   1410
      Width           =   750
   End
   Begin VB.TextBox txtDes_General 
      Height          =   795
      Left            =   1530
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   2325
      Width           =   4635
   End
   Begin VB.TextBox txtCod_DivPreLOT 
      Height          =   300
      Left            =   1530
      MaxLength       =   3
      TabIndex        =   13
      Top             =   1995
      Width           =   630
   End
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
      TabIndex        =   11
      Text            =   "0"
      Top             =   1680
      Width           =   750
   End
   Begin VB.Frame fraNORegular 
      Caption         =   "Datos P.O.No Regular"
      Height          =   1440
      Left            =   30
      TabIndex        =   6
      Top             =   3120
      Width           =   6210
      Begin VB.TextBox txtPrecio_RecCliLOT 
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
         Left            =   1590
         TabIndex        =   7
         Text            =   "0"
         Top             =   315
         Width           =   750
      End
      Begin MSComCtl2.DTPicker dtpFec_RecCliLOT 
         Height          =   315
         Left            =   1590
         TabIndex        =   8
         Top             =   750
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   16842753
         CurrentDate     =   37159
      End
      Begin VB.Label labels 
         Caption         =   "Precio del Cliente"
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   10
         Tag             =   "Client Price"
         Top             =   345
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "Fecha Ingreso a Almacén"
         Height          =   450
         Index           =   22
         Left            =   135
         TabIndex        =   9
         Tag             =   "Reception Warehouse Date"
         Top             =   735
         Width           =   1440
      End
   End
   Begin VB.TextBox txtDes_DestinoLOT 
      Height          =   285
      Left            =   2190
      MaxLength       =   30
      TabIndex        =   3
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
   Begin MSComCtl2.DTPicker dtpFec_DespachoOriLOT 
      Height          =   315
      Left            =   1530
      TabIndex        =   1
      Top             =   705
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16842753
      CurrentDate     =   37159
   End
   Begin FunctionsButtons.FunctButt acbForm 
      Height          =   510
      Left            =   1905
      TabIndex        =   2
      Top             =   5400
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmUpdateDatGenLotEst.frx":0000
      Orientacion     =   0
      Style           =   1
      Language        =   1
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Label labels 
      Caption         =   "Utilidad Cotizada"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   26
      Left            =   3120
      TabIndex        =   33
      Tag             =   "Price"
      Top             =   1965
      Width           =   1335
   End
   Begin VB.Label labels 
      Caption         =   "Precio Cotizado"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   25
      Left            =   3120
      TabIndex        =   31
      Tag             =   "Price"
      Top             =   1695
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Lead Time :"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Pre Pack :"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label labels 
      Caption         =   "% Comisión"
      Height          =   255
      Index           =   18
      Left            =   3150
      TabIndex        =   23
      Tag             =   "Commision"
      Top             =   1110
      Width           =   1335
   End
   Begin VB.Label labels 
      Caption         =   "Modo de Comisión"
      Height          =   255
      Index           =   20
      Left            =   60
      TabIndex        =   21
      Top             =   1065
      Width           =   1305
   End
   Begin VB.Label labels 
      Caption         =   "Importe Comisión"
      Height          =   240
      Index           =   1
      Left            =   3150
      TabIndex        =   20
      Top             =   1410
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Letter Credit :"
      Height          =   195
      Left            =   60
      TabIndex        =   16
      Tag             =   "Letter Credit :"
      Top             =   2325
      Width           =   945
   End
   Begin VB.Label labels 
      Caption         =   "Division de Prenda"
      Height          =   255
      Index           =   19
      Left            =   60
      TabIndex        =   14
      Tag             =   "Garment Division"
      Top             =   1980
      Width           =   1335
   End
   Begin VB.Label labels 
      Caption         =   "Precio"
      Height          =   255
      Index           =   14
      Left            =   60
      TabIndex        =   12
      Tag             =   "Price"
      Top             =   1695
      Width           =   1335
   End
   Begin VB.Label labels 
      Caption         =   "Destino"
      Height          =   255
      Index           =   15
      Left            =   60
      TabIndex        =   5
      Tag             =   "Destination"
      Top             =   75
      Width           =   1200
   End
   Begin VB.Label labels 
      Caption         =   "Ex-Factory"
      Height          =   255
      Index           =   17
      Left            =   60
      TabIndex        =   4
      Tag             =   "Delivery Date"
      Top             =   765
      Width           =   1320
   End
End
Attribute VB_Name = "frmUpdateDatGenLotEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oParent         As Object

Public sCod_PurOrd     As String

Public sCod_LotPurOrd  As String

Public sCod_Cliente    As String

Public sCod_EstCli     As String

Public sFlag           As String

Public sCod_Destino    As String

Public sCod_DestinoLOT As String

Public sFlg_Regular    As String

Dim Rs_cargaanexo      As ADODB.Recordset

Public Codigo, Descripcion As String

Public Sub CARGA_DATOSANEXOS()

    Dim strSql As String
    
    Set Rs_cargaanexo = New ADODB.Recordset
    Rs_cargaanexo.ActiveConnection = cCONNECT
    Rs_cargaanexo.CursorType = adOpenStatic
    Rs_cargaanexo.CursorLocation = adUseClient
    Rs_cargaanexo.LockType = adLockReadOnly
        
    'Esta cadena es la que nos devolvera los grupos de produccion
    strSql = "SELECT Cod_DivPre, Precio_RecCli, Fec_RecCli FROM TG_LOTEST WHERE " & "Cod_Cliente ='" & sCod_Cliente & "' AND " & "Cod_PurOrd ='" & sCod_PurOrd & "' AND " & "Cod_LotPurOrd='" & sCod_LotPurOrd & "' AND " & "Cod_EstCli='" & sCod_EstCli & "'"

    Rs_cargaanexo.Open strSql

    If Rs_cargaanexo.RecordCount > 0 Then
        If Not IsNull(Rs_cargaanexo("Cod_DivPre").value) Then
            txtCod_DivPreLOT.Text = Rs_cargaanexo("Cod_DivPre").value
        Else
            txtCod_DivPreLOT.Text = ""
        End If
        
        If Not IsNull(Rs_cargaanexo("Precio_RecCli").value) Then
            txtPrecio_RecCliLOT.Text = Rs_cargaanexo("Precio_RecCli").value
        Else
            txtPrecio_RecCliLOT.Text = ""
        End If

        If Not IsNull(Rs_cargaanexo("Fec_RecCli")) Then
            dtpFec_RecCliLOT.value = Rs_cargaanexo("Fec_RecCli").value
        Else
            dtpFec_RecCliLOT.value = Date
        End If
        
    End If
    
    Rs_cargaanexo.Close
    Set Rs_cargaanexo = Nothing
    
End Sub

Private Sub acbForm_ActionClick(ByVal Index As Integer, _
                                ByVal ActionType As Integer, _
                                ByVal ActionName As String)

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

    Dim sFlg_NoRegular

    Dim sComisionEnPorcentaje As String
    
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
    
    If optComisionEnPorcentaje Then
        sComisionEnPorcentaje = "S"
    Else
        sComisionEnPorcentaje = "N"
    End If
        
    'objPO.UpdateDatGenPurORd sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, txtCod_DestinoLOT.Text, FechaOK(dtpFec_DespachoActLOT.value), CDbl(txtPor_ComisionLOT.Text), vusu, ComputerName, FechaOK(dtpFec_DespachoOriLOT.value), FixNulos(txtPrecioLOT.Text, vbDouble), sFlg_Regular, FixNulos(txtPrecio_RecCliLOT.Text, vbDouble), FechaOK(dtpFec_RecCliLOT.value), txtCod_DivPreLOT.Text, txtDes_General.Text, sComisionEnPorcentaje, CDbl(txtImp_Comision.Text)
    objPO.UpdateDatGenPurORd sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, txtCod_DestinoLOT.Text, CDbl(txtPor_ComisionLOT.Text), vusu, ComputerName, FechaOK(dtpFec_DespachoOriLOT.value), FixNulos(txtPrecioLOT.Text, vbDouble), sFlg_Regular, FixNulos(txtPrecio_RecCliLOT.Text, vbDouble), FechaOK(dtpFec_RecCliLOT.value), txtCod_DivPreLOT.Text, txtDes_General.Text, sComisionEnPorcentaje, CDbl(txtImp_Comision.Text), Trim(Txtcod_Lead.Text)
    '    oParent.Buscar
    '    oParent.BuscarEStilos
    
    Set objPO = Nothing
    Unload Me

    Exit Function

errores:

    If Not objPO Is Nothing Then
        Set objPO = Nothing
    End If
    
    ErrorHandler Err, Err.Description

End Function

Private Function VAlidFechaDespacho(dFecha As String) As Boolean

    On Error GoTo errores

    Dim vbuff

    Dim obj  As clsTG_LotColTal

    Dim iRet As Integer
    
    Set obj = New clsTG_LotColTal
    obj.ConexionString = cCONNECT
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
    
    ErrorHandler Err, Err.Description
End Function

Private Sub txtCod_DestinoLOT_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        sFlag = "COD_DESTINOLOT"

        If Filtrar(sFlag, Me, txtCod_DestinoLOT, txtDes_DestinoLOT) Then
            Me.dtpFec_DespachoOriLOT.SetFocus
        End If
    End If

End Sub

Public Function ValidStep() As Boolean

    Dim aMess(4)

    Dim amensaje As clsMessages

    Set amensaje = New clsMessages
  
    If txtCod_DestinoLOT.Text = "" Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY

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

    If dtpFec_DespachoOriLOT.value = "" Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY

        If dtpFec_DespachoOriLOT.Enabled Then
            Me.dtpFec_DespachoOriLOT.SetFocus
        End If

        Exit Function

    End If
    
    If Not VAlidFechaDespacho(FechaOK(dtpFec_DespachoOriLOT.value)) Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_INVALID_SELECC

        If dtpFec_DespachoOriLOT.Enabled Then
            dtpFec_DespachoOriLOT.SetFocus
        End If

        Exit Function

    End If
    
    '    If Not VAlidFechaDespacho(FechaOK(dtpFec_DespachoOriLOT.Value)) Then
    '        Mensaje MESSAGECODE.kMESSAGE_ERR_INVALID_SELECC
    '        If dtpFec_DespachoOriLOT.Enabled Then
    '            dtpFec_DespachoOriLOT.SetFocus
    '        End If
    '        Exit Function
    '    End If
    
    '    If optComisionEnPorcentaje And CDbl(txtPor_ComisionLOT.Text) <= 0 Then
    '        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
    '        If txtPor_ComisionLOT.Enabled Then
    '            txtPor_ComisionLOT.SetFocus
    '            Exit Function
    '        End If
    '    End If
    '
    '    If optComisionEnImporte And CDbl(txtImp_Comision.Text) <= 0 Then
    '        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
    '        If txtImp_Comision.Enabled Then
    '            txtImp_Comision.SetFocus
    '            Exit Function
    '        End If
    '    End If
    '
    
    If FixNulos(txtPrecioLOT.Text, vbDouble) = 0 Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY

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
    
    If sFlg_Regular = "N" Then
        If FixNulos(txtPrecio_RecCliLOT.Text, vbDouble) = 0 Then
            Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY

            If txtPrecio_RecCliLOT.Enabled Then
                Me.txtPrecio_RecCliLOT.SetFocus
            End If

            Exit Function

        End If
    End If
    
    ValidStep = True
End Function

Private Function ValidCod_DestinoLot() As Boolean

    sFlag = "COD_DESTINO"

    If Not Filtrar(sFlag, Me, Me.txtCod_DestinoLOT, Me.txtDes_DestinoLOT, False) Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTFOUND

        If Me.txtCod_DestinoLOT.Enabled Then
            Me.txtCod_DestinoLOT.SetFocus
        End If

        Exit Function

    End If

    ValidCod_DestinoLot = True
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

Private Function VAlidDivPre(sCod_DivPRe As String) As Boolean

    On Error GoTo errores

    Dim vbuff

    Dim obj    As clsTG_LotColTal

    Dim bValid As Boolean
    
    Set obj = New clsTG_LotColTal
    obj.ConexionString = cCONNECT
    bValid = obj.VAlidDivPre(sCod_DivPRe)
    Set obj = Nothing
    
    If Not bValid Then
        Load frmDivPre
        Set frmDivPre.oParent = Me
        frmDivPre.sCod_DivPRe = Me.txtCod_DivPreLOT.Text
        frmDivPre.txtCod_DivPre.Text = frmDivPre.sCod_DivPRe
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
    
    ErrorHandler Err, Err.Description
End Function

Private Sub optComisionEnImporte_Click()
    txtPor_ComisionLOT.Enabled = False
    txtImp_Comision.Enabled = True
    txtPor_ComisionLOT.Text = 0

    If Me.txtImp_Comision.Enabled Then
        If Me.Visible Then
            txtImp_Comision.SetFocus
        End If
    End If

End Sub

Private Sub optComisionEnPorcentaje_Click()
    txtPor_ComisionLOT.Enabled = True
    
    txtImp_Comision.Text = 0
    txtImp_Comision.Enabled = False

    If Me.txtPor_ComisionLOT.Enabled Then
        txtPor_ComisionLOT.SetFocus

    End If

End Sub

Private Sub Txtcod_Lead_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then Busca_Lead_Time 1
End Sub

Private Sub Busca_Lead_Time(opcion As Integer)

    Dim sField As String, iRows As Long

    Dim rstAux As ADODB.Recordset

    Dim strSql As String

    Select Case opcion

        Case 1: strSql = "SELECT Tip_Lead_Time, Des_Tip_Lead_Time FROM TG_Lead_Time where Tip_Lead_Time like '%" & Trim(Txtcod_Lead) & "%'"

        Case 2: strSql = "SELECT Tip_Lead_Time, Des_Tip_Lead_Time FROM TG_Lead_Time where Des_Tip_Lead_Time like '%" & Trim(TxtLead) & "%'"
    End Select
    
    Txtcod_Lead = "": TxtLead = ""

    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = strSql
        .Caption = "Seleccionar Tipos Lead Time"
        .Cargar_Datos
        Codigo = "": Descripcion = ""
        
        iRows = .gexLista.RowCount
        Set rstAux = .gexLista.ADORecordset

        If .gexLista.RowCount > 1 Then
            .Show vbModal
        ElseIf .gexLista.RowCount = 1 Then
            Codigo = .gexLista.value(.gexLista.Columns("Tip_Lead_Time").Index)
            Descripcion = .gexLista.value(.gexLista.Columns("Des_Tip_Lead_Time").Index)
        End If
        
        If Codigo <> "" Then
            Txtcod_Lead = RTrim(Codigo)
            TxtLead = RTrim(Descripcion)
        End If

    End With

    ''''txtPOrcentaje.SetFocus
    acbForm.SetFocus
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub

Private Sub TxtLead_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then Busca_Lead_Time 2
End Sub

Private Sub txtPor_ComisionLOT_GotFocus()
    SelectionText txtPor_ComisionLOT
End Sub

Private Sub txtPor_ComisionLOT_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If optComisionEnPorcentaje Then
            txtImp_Comision.Text = 0
        End If
    End If

End Sub

Private Sub txtImp_Comision_GotFocus()
    SelectionText txtImp_Comision
End Sub

Private Sub txtImp_Comision_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And optComisionEnImporte.value Then
        txtImp_Comision.Text = FixNulos(CDbl(txtImp_Comision.Text), vbDouble)
        txtPor_ComisionLOT.Text = 0
    End If

End Sub

Private Sub txtPrecioLOT_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        txtCod_DivPreLOT.SetFocus
    End If

End Sub
