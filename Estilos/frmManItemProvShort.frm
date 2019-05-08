VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "mantenimientos.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmManItemProvShort 
   Caption         =   "Items del Proveedor"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDetalle 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   60
      TabIndex        =   13
      Top             =   3105
      Width           =   7350
      Begin VB.TextBox txtDes_Proveedor 
         Height          =   285
         Left            =   2910
         TabIndex        =   2
         Top             =   270
         Width           =   4260
      End
      Begin VB.TextBox txtCod_Proveedor 
         Height          =   285
         Left            =   1300
         MaxLength       =   12
         TabIndex        =   1
         Top             =   270
         Width           =   1560
      End
      Begin VB.TextBox txtCod_ItemProv 
         Height          =   285
         Left            =   1300
         MaxLength       =   10
         TabIndex        =   3
         Top             =   600
         Width           =   1545
      End
      Begin VB.TextBox txtCod_UniMedProv 
         Height          =   285
         Left            =   4350
         MaxLength       =   4
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtPrecioCotizado 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1305
         TabIndex        =   5
         Text            =   "0"
         Top             =   945
         Width           =   1545
      End
      Begin VB.TextBox txtObservacioes 
         Height          =   315
         Left            =   1305
         TabIndex        =   6
         Top             =   1335
         Width           =   5850
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "U.M. Proveedor :"
         Height          =   195
         Left            =   2940
         TabIndex        =   18
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Item Prov:"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor :"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   330
         Width           =   825
      End
      Begin VB.Label Label7 
         Caption         =   "Precio Cotizado  :"
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   1035
         Width           =   1125
      End
      Begin VB.Label Label8 
         Caption         =   "Observaciones   :"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1380
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   105
      TabIndex        =   8
      Top             =   5010
      Width           =   1965
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   15
         Picture         =   "frmManItemProvShort.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   495
         Picture         =   "frmManItemProvShort.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   975
         Picture         =   "frmManItemProvShort.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmManItemProvShort.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame FraLista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   7380
      Begin VB.Frame fraCambio 
         Caption         =   "Cambio de Código Item"
         Height          =   1575
         Left            =   1515
         TabIndex        =   21
         Top             =   1155
         Visible         =   0   'False
         Width           =   3855
         Begin VB.TextBox txtNuevoCodigo 
            Height          =   315
            Left            =   1275
            TabIndex        =   22
            Top             =   405
            Width           =   2385
         End
         Begin FunctionsButtons.FunctButt FunctButt2 
            Height          =   510
            Left            =   750
            TabIndex        =   24
            Top             =   945
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   900
            Custom          =   "0~0~ACEPTAR~True~True~&Aceptar~0~0~1~~0~False~False~&Aceptar~~1~0~CANCELAR~True~True~Cancelar~1~0~3~~0~False~False~Cancelar~"
            Orientacion     =   0
            Style           =   0
            Language        =   0
            TypeImageList   =   0
            ControlWidth    =   1155
            ControlHeigth   =   490
            ControlSeparator=   110
         End
         Begin VB.Label Label1 
            Caption         =   "Nuevo Codigo"
            Height          =   285
            Left            =   150
            TabIndex        =   23
            Top             =   450
            Width           =   1170
         End
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2655
         Left            =   90
         TabIndex        =   19
         Top             =   225
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   4683
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         Enabled         =   0   'False
         MethodHoldFields=   -1  'True
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmManItemProvShort.frx":05C8
         Column(2)       =   "frmManItemProvShort.frx":0690
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmManItemProvShort.frx":0734
         FormatStyle(2)  =   "frmManItemProvShort.frx":086C
         FormatStyle(3)  =   "frmManItemProvShort.frx":091C
         FormatStyle(4)  =   "frmManItemProvShort.frx":09D0
         FormatStyle(5)  =   "frmManItemProvShort.frx":0AA8
         FormatStyle(6)  =   "frmManItemProvShort.frx":0B60
         ImageCount      =   0
         PrinterProperties=   "frmManItemProvShort.frx":0C40
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2310
      TabIndex        =   7
      Top             =   5085
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmManItemProvShort.frx":0E18
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   525
      Left            =   6075
      TabIndex        =   20
      Top             =   5115
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~CAMBIO~True~True~&Cambio Codigo Item~0~0~1~~0~False~False~&Cambio Codigo Item~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmManItemProvShort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs_Lista As ADODB.Recordset
Dim strSQL As String
Dim sTipo As String
Public varCod_item As String, varCod_Proveedor As String
Public Codigo As String, Descripcion As String
Public sUniMedDefault As String

Private Sub Form_Load()
    'Call FormateaGrid(DGridLista)
    Call INHABILITA_DATOS
    Me.MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    Me.FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub
Private Sub cmdFirst_Click()
'    If Not Rs_Lista.BOF Then
'        Rs_Lista.MoveFirst
'    End If
If Not GridEX1.ADORecordset.BOF Then
    GridEX1.MoveFirst
End If
End Sub

Private Sub cmdLast_Click()
'    If Not Rs_Lista.EOF Then
'        Rs_Lista.MoveLast
'    End If
If Not GridEX1.ADORecordset.BOF Then
    GridEX1.MoveLast
End If

End Sub

Private Sub cmdNext_Click()
'    If Not Rs_Lista.EOF Then
'        Rs_Lista.MoveNext
'        If Rs_Lista.EOF Then
'            Rs_Lista.MoveLast
'        End If
'     end if
    If Not GridEX1.ADORecordset.BOF Then
        'GridEX1.ADORecordset.MoveNext
        GridEX1.MoveNext
        If GridEX1.ADORecordset.EOF Then
            GridEX1.MoveLast
        End If
    End If
End Sub

Private Sub cmdPrevious_Click()
'    If Not Rs_Lista.BOF Then
'        Rs_Lista.MovePrevious
'        If Rs_Lista.BOF Then
'            Rs_Lista.MoveFirst
'        End If
'    End If
    If Not GridEX1.ADORecordset.BOF Then
        GridEX1.MovePrevious
        If GridEX1.ADORecordset.BOF Then
            GridEX1.MoveFirst
        End If
    End If
End Sub

Sub LIMPIAR_DATOS()

    txtCod_Proveedor.Text = ""
    txtDes_Proveedor.Text = ""
    txtCod_ItemProv.Text = ""

    txtCod_UniMedProv = ""
    txtPrecioCotizado.Text = "0.00"
    txtObservacioes.Text = ""
    
End Sub


Function VALIDA_DATOS() As Boolean
    Dim NombreTabla As String
    Dim CodigoTabla As String
    

    VALIDA_DATOS = True
    If sTipo <> "D" Then

        If sTipo = "I" Then
        
            strSQL = "SELECT COUNT(*) FROM LG_ITEMPROV WHERE Cod_Item='" & varCod_item & "' AND Cod_Proveedor='" & Trim(txtCod_Proveedor.Text) & "' AND Cod_ItemProv='" & Trim(txtCod_ItemProv.Text) & "'"
            
            If DevuelveCampo(strSQL, cCONNECT) <> "0" Then
                MsgBox "El código de item de proveedor ya se encuentra registrado. Sirvase verificar", vbInformation, "Item Proveedor"
                txtCod_ItemProv.SetFocus
                VALIDA_DATOS = False
                Exit Function
            End If
            
            'If Trim(txtCod_ItemProv.Text) = "" Then
            '    MsgBox "El Código de Item de Proveedor no puede estar vacio. Sirvase verificar", vbInformation, "Item Proveedor"
            '    txtCod_ItemProv.Text = ""
            '    txtCod_ItemProv.SetFocus
            '    VALIDA_DATOS = False
            'Exit Function
            
            'End If
            
        End If

'        If Trim(txtcod_StaOrdComp.Text) = "" Then
'            MsgBox "El código de Status de Orden de Compra no puede estar vacío. Sirvase verificar", vbInformation, "Ordenes de Compra"
'            txtcod_StaOrdComp.Text = ""
'            txtcod_StaOrdComp.SetFocus
'            VALIDA_DATOS = False
'            Exit Function
'        End If
'
'        If Trim(txtDes_StaOrdComp.Text) = "" Then
'            MsgBox "La descripción de Status de Orden de Compra no puede estar vacío. Sirvase verificar", vbInformation, "Ordenes de Compra"
'            txtDes_StaOrdComp.Text = ""
'            txtDes_StaOrdComp.SetFocus
'            VALIDA_DATOS = False
'            Exit Function
'        End If

        If Trim(txtCod_Proveedor.Text) = "" Then
            MsgBox "El Código de Proveedor no puede estar vacio. Sirvase verificar", vbInformation, "Item Proveedor"
            txtCod_Proveedor.Text = ""
            txtCod_Proveedor.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        
        strSQL = "SELECT count(*) FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & txtCod_Proveedor.Text & "'"
        If DevuelveCampo(strSQL, cCONNECT) = "0" Then
            MsgBox "El código de proveedor ingresado no es válido. Sirvase verificar", vbInformation, "Item Proveedor"
            txtCod_Proveedor.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If

        If Trim(txtCod_UniMedProv.Text) = "" Then
            MsgBox "La unidad de medida no puede estar vacia. Sirvase verificar", vbInformation, "Item Proveedor"
            txtCod_UniMedProv.Text = ""
            txtCod_UniMedProv.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If

        'txtFac_EquiProv

    Else
'        'Aqui se valida que no tenga registros dependientes
'        Strsql = "SELECT COUNT(*) FROM LG_ORDCOMPITEM WHERE Ser_OrdComp='" & Rs_Lista("Ser_OrdComp").Value & "' AND Cod_OrdComp='" & Rs_Lista("Cod_OrdComp").Value & "'"
'        If DevuelveCampo(Strsql, cCONNECT) > 0 Then
'            MsgBox "El registro seleccionado posee registros relacionados. Sirvase verificar", vbInformation, "Ordenes de Compra"
'            VALIDA_DATOS = False
'            Exit Function
'        End If
    End If
End Function

Sub Carga_Datos()

'    If Not Rs_Lista.EOF And Not Rs_Lista.BOF Then
'
'        txtCod_Proveedor.Text = Rs_Lista("Cod_Proveedor").Value
'        txtDes_Proveedor.Text = Rs_Lista("Des_Proveedor").Value
'        txtCod_ItemProv.Text = Rs_Lista("Cod_ItemProv").Value
'        txtCod_UniMedProv.Text = Rs_Lista("Cod_UniMedProv").Value
'        txtPrecioCotizado.Text = Rs_Lista("Pre_Cotizado").Value
'        txtObservacioes.Text = Rs_Lista("Observaciones").Value
'   End If
''''
    If Not GridEX1.RowCount > 0 Then Exit Sub
    txtCod_Proveedor.Text = GridEX1.Value(GridEX1.Columns("Cod_Proveedor").Index)
    txtDes_Proveedor.Text = GridEX1.Value(GridEX1.Columns("Des_Proveedor").Index)
    txtCod_ItemProv.Text = GridEX1.Value(GridEX1.Columns("Cod_ItemProv").Index)
    txtCod_UniMedProv.Text = GridEX1.Value(GridEX1.Columns("Cod_UniMedProv").Index)
    txtPrecioCotizado.Text = GridEX1.Value(GridEX1.Columns("Pre_Cotizado").Index)
    txtObservacioes.Text = GridEX1.Value(GridEX1.Columns("Observaciones").Index)
''''

End Sub

Sub HABILITA_DATOS()
    
    txtCod_UniMedProv.Enabled = True
    txtPrecioCotizado.Enabled = True
    txtObservacioes.Enabled = True
    
    If sTipo = "I" Then
        txtCod_Proveedor.Enabled = True
        txtDes_Proveedor.Enabled = True
        txtCod_ItemProv.Enabled = True
        txtPrecioCotizado.Enabled = True
        txtObservacioes.Enabled = True
        txtCod_Proveedor.SetFocus
    Else
        txtCod_UniMedProv.SetFocus
    End If
    
End Sub

Sub INHABILITA_DATOS()
    
    txtCod_Proveedor.Enabled = False
    txtDes_Proveedor.Enabled = False
    txtCod_ItemProv.Enabled = False
    txtCod_UniMedProv.Enabled = False
    txtPrecioCotizado.Enabled = False
    txtObservacioes.Enabled = False
    
    
End Sub

Sub CARGA_GRID()
'    Dim xRow As Variant
'
'    xRow = DGridLista.Row
'
'    Set Rs_Lista = New ADODB.Recordset
'    Rs_Lista.ActiveConnection = cCONNECT
'    Rs_Lista.CursorType = adOpenStatic
'    Rs_Lista.CursorLocation = adUseClient
'    Rs_Lista.LockType = adLockReadOnly
'
'    'Esta cadena es para devolver el Codigo de Cliente
'    StrSQL = "EXEC UP_SEL_ITEMPROV '" & varCod_item & "'" '& varCod_Proveedor & "'"
'
'
'
'    Rs_Lista.Open StrSQL
'    Set DGridLista.DataSource = Rs_Lista
'    DGridLista.Refresh
'
'    If xRow > 0 And xRow <= Rs_Lista.RecordCount Then
'        'DGridLista.Row = xRow
'        DGridLista.Row = 0
'    End If
'
'    If Rs_Lista.RecordCount > 0 Then
'        DGridLista.Enabled = True
'        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
'        Call Carga_Datos
'    Else
'        DGridLista.Enabled = False
'        HabilitaMant Me.MantFunc1, "ADICIONAR"
'        Call LIMPIAR_DATOS
'    End If
'
'      Me.MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
'''''
    'StrSQL = "EXEC UP_SEL_ITEMPROV '1','" & Me.DTPicker1 & "','" & Me.DTPicker1 & "','" & Me.txtCod_Modulo & "'"
    strSQL = "EXEC UP_SEL_ITEMPROV '" & varCod_item & "'" '& varCod_Proveedor & "'"

    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
    
    GridEX1.Columns("Des_Proveedor").Width = 2200
    GridEX1.Columns("Cod_ItemProv").Width = 1200
    GridEX1.Columns("Cod_UniMedProv").Width = 600
    GridEX1.Columns("Pre_Cotizado").Width = 1400
    GridEX1.Columns("Observaciones").Width = 1800
    
    GridEX1.Columns("Des_Proveedor").Caption = "Proveedor"
    GridEX1.Columns("Cod_ItemProv").Caption = "Item Prov"
    GridEX1.Columns("Cod_UniMedProv").Caption = "U.M."
    GridEX1.Columns("Pre_Cotizado").Caption = "Precio Cotizado"
    GridEX1.Columns("Observaciones").Caption = "Observaciones"
    GridEX1.Columns("Cod_Item").Width = 0
    GridEX1.Columns("Cod_Proveedor").Width = 0
    GridEX1.Columns("Fac_EquiProv").Width = 0
    GridEX1.Columns("Precio").Width = 0
    If GridEX1.ADORecordset.RecordCount > 0 Then
        GridEX1.Enabled = True
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Call Carga_Datos
    Else
        GridEX1.Enabled = False
        HabilitaMant Me.MantFunc1, "ADICIONAR"
        Call LIMPIAR_DATOS
    End If
    
'''''
End Sub

Function SALVAR_DATOS() As Boolean
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim strSQL As String
    
    Con.ConnectionString = cCONNECT
    Con.Open
    
        Con.BeginTrans

        strSQL = "EXEC UP_MAN_ITEMPROV_SHORT '" & _
        sTipo & "','" & _
        varCod_item & "','" & _
        Trim(txtCod_Proveedor.Text) & "','" & _
        Trim(txtCod_ItemProv.Text) & "','" & _
        Trim(txtCod_UniMedProv.Text) & "'," & _
        txtPrecioCotizado.Text & ",'" & _
        txtObservacioes.Text & "'"

        
        
        Con.Execute strSQL

        Con.CommitTrans
        Dim amensaje As New clsMessages
        amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
        Informa "", amensaje
       
        SALVAR_DATOS = True
        
    Exit Function
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Function
Sub ELIMINAR_DATOS()
 Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
   
    Con.ConnectionString = cCONNECT
    Con.Open
    Con.BeginTrans
       
        strSQL = "EXEC UP_MAN_ITEMPROV_SHORT'" & _
        sTipo & "','" & _
        varCod_item & "','" & _
        Trim(txtCod_Proveedor.Text) & "','" & _
        Trim(txtCod_ItemProv.Text) & "','" & _
        Trim(txtCod_UniMedProv.Text) & "',0,0"
        
        Con.Execute strSQL
    
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje
    
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"

End Sub

Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call Carga_Datos
End Sub

Sub BUSCA_PROVEEDOR(tipo As Integer)
    Select Case tipo
        Case 1:
                
                strSQL = "SELECT Des_Proveedor FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & txtCod_Proveedor.Text & "'"
                txtDes_Proveedor.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                'Strsql = "SELECT Cod_Proveedor FROM LG_PROVEEDOR WHERE Des_Proveedor = '" & txtDes_Proveedor.Text & "'"
                'txtCod_Proveedor.Text = Trim(DevuelveCampo(Strsql, cCONNECT))
                txtCod_ItemProv.SetFocus

                
        Case 2:
                Dim oTipo As New frmBusqGeneral
                Dim rs As New ADODB.Recordset
                Set oTipo.oParent = Me
                oTipo.sQuery = "SELECT Cod_Proveedor as Código, Des_Proveedor as Descripción FROM LG_PROVEEDOR WHERE Des_Proveedor like '%" & Trim(txtDes_Proveedor.Text) & "%'"
                oTipo.Cargar_Datos
                oTipo.Show 1
                If Codigo <> "" Then
                    txtCod_Proveedor.Text = Trim(Codigo)
                    txtDes_Proveedor.Text = Trim(Descripcion)
                    Codigo = ""
                    Descripcion = ""
                    txtCod_ItemProv.SetFocus
                End If
                Set oTipo = Nothing
                Set rs = Nothing
                
    End Select
End Sub





Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    fraCambio.Visible = True
    txtNuevoCodigo.Text = ""
    txtNuevoCodigo.SetFocus
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call CambioCodigo
Case "CANCELAR"
    fraCambio.Visible = False
End Select

End Sub

Private Sub GridEX1_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Call Carga_Datos
End Sub

Sub CambioCodigo()

On Error GoTo hand
    
        If MsgBox("Esta seguro de cambiar el código", vbInformation + vbYesNo, "INFORMACIÓN") = vbYes Then
          

          strSQL = "EXEC LG_CAMBIA_CODIGO_ITEM_PROVEEDOR '" & Trim(GridEX1.Value(GridEX1.Columns("Cod_Item").Index)) & "','" & Trim(txtCod_Proveedor.Text) & "','" & Trim(txtCod_ItemProv.Text) & "','" & Trim(txtNuevoCodigo.Text) & "'"
         
            If ExecuteSQL(cCONNECT, strSQL) = -1 Then
         
                MsgBox "El cambio de código se realizo exitosamente", vbInformation, "Cierre"
                txtNuevoCodigo.Text = ""
                fraCambio.Visible = False
            End If
          
        End If
    CARGA_GRID
    Exit Sub
hand:
    ErrorHandler Err, "CAMBIO_CODIGO_ITEM"
 
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
    Dim fnuevo As Boolean
    Dim ItemsProvACambiar As String
    
    
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            LIMPIAR_DATOS
            If sUniMedDefault <> "" Then
                txtCod_UniMedProv.Text = sUniMedDefault
            End If
            HABILITA_DATOS
            'txtCod_Proveedor.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            'DGridLista.Enabled = False
        Case "MODIFICAR"
        
            sTipo = "U"
            HABILITA_DATOS
            'txtCod_Proveedor.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            'DGridLista.Enabled = False
        Case "ELIMINAR"
        
            Eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Item Proveedor")
            If Eliminar = vbYes Then
                sTipo = "D"
                If VALIDA_DATOS Then
                    Call ELIMINAR_DATOS
                    Call CARGA_GRID
                    sTipo = ""
                End If
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                If sTipo = "U" Then ItemsProvACambiar = GridEX1.Value(GridEX1.Columns("Cod_ItemProv").Index)
                If SALVAR_DATOS Then
                    Call CARGA_GRID
                    Call INHABILITA_DATOS
                  HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                    'DGridLista.Enabled = True
                    If sTipo = "I" Then
                        MantFunc1_ActionClick 0, 0, "ADICIONAR"
                    End If
                    If sTipo = "U" Then
                        fnuevo = GridEX1.Find(GridEX1.Columns("Cod_ItemProv").Index, jgexGreaterThanOrEqualTo, ItemsProvACambiar)
                    End If
                    
                End If
            End If
        Case "DESHACER"
            Call LIMPIAR_DATOS
            Call Carga_Datos
            Call INHABILITA_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            'DGridLista.Enabled = True
            sTipo = ""
         Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub txtCod_ItemProv_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtCod_UniMedProv.SetFocus
    End If
End Sub

Private Sub txtCod_Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Proveedor.Text) <> "" Then
            Call BUSCA_PROVEEDOR(1)
        End If
    End If
End Sub

Private Sub txtCod_UniMedProv_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtPrecioCotizado.SetFocus
    End If
End Sub

Private Sub txtDes_Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDes_Proveedor.Text) <> "" Then
            Call BUSCA_PROVEEDOR(2)
        End If
    End If
End Sub

Private Sub txtFac_EquiProv_KeyPress(KeyAscii As Integer)
    'Call SoloNumeros(txtFac_EquiProv, KeyAscii, True, 6, 7)
    If KeyAscii = vbKeyReturn Then
        txtCod_UniMedProv.SetFocus
    End If
End Sub


Private Sub txtLeadTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtPrecioCotizado.SetFocus
    End If
End Sub


Private Sub txtNuevoCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FunctButt2.SetFocus
    End If
End Sub

Private Sub txtObservacioes_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        MantFunc1.SetFocus
    End If
End Sub



Private Sub txtPrecioCotizado_GotFocus()
    SelectionText txtPrecioCotizado
End Sub

Private Sub txtPrecioCotizado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtObservacioes.SetFocus
    End If
End Sub


