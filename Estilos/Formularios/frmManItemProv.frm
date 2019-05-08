VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManItemProv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Proveedor"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2850
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6210
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2565
         Left            =   60
         TabIndex        =   17
         Top             =   195
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4524
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "Des_Proveedor"
            Caption         =   "Proveedor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Cod_ItemProv"
            Caption         =   "Item Prov."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Fac_EquiProv"
            Caption         =   "F. Equiv."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Cod_UniMedProv"
            Caption         =   "U.M."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Precio"
            Caption         =   "Precio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Pre_Cotizado"
            Caption         =   "Precio Cotizado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Observaciones"
            Caption         =   "Observaciones"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   2
            BeginProperty Column00 
               ColumnWidth     =   2069.858
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   240
      TabIndex        =   11
      Top             =   5625
      Width           =   1965
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmManItemProv.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   975
         Picture         =   "frmManItemProv.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   495
         Picture         =   "frmManItemProv.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   15
         Picture         =   "frmManItemProv.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
   End
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
      Height          =   2670
      Left            =   0
      TabIndex        =   1
      Top             =   2865
      Width           =   6210
      Begin VB.TextBox txtObservacioes 
         Height          =   315
         Left            =   1560
         TabIndex        =   25
         Top             =   1920
         Width           =   4335
      End
      Begin VB.TextBox txtPrecioCotizado 
         Height          =   315
         Left            =   1560
         TabIndex        =   23
         Text            =   "0"
         Top             =   1590
         Width           =   1695
      End
      Begin VB.TextBox txtLeadTime 
         Height          =   315
         Left            =   4560
         TabIndex        =   21
         Text            =   "0"
         Top             =   1260
         Width           =   1245
      End
      Begin VB.TextBox txtPrecio 
         Height          =   315
         Left            =   1320
         TabIndex        =   20
         Text            =   "0"
         Top             =   1260
         Width           =   1695
      End
      Begin VB.TextBox txtCod_UniMedProv 
         Height          =   285
         Left            =   4560
         MaxLength       =   4
         TabIndex        =   10
         Top             =   930
         Width           =   1245
      End
      Begin VB.TextBox txtFac_EquiProv 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1320
         TabIndex        =   8
         Text            =   "1.00"
         Top             =   930
         Width           =   1680
      End
      Begin VB.TextBox txtCod_ItemProv 
         Height          =   285
         Left            =   1300
         MaxLength       =   10
         TabIndex        =   6
         Top             =   600
         Width           =   1680
      End
      Begin VB.TextBox txtCod_Proveedor 
         Height          =   285
         Left            =   1300
         MaxLength       =   12
         TabIndex        =   3
         Top             =   270
         Width           =   1305
      End
      Begin VB.TextBox txtDes_Proveedor 
         Height          =   285
         Left            =   2700
         TabIndex        =   4
         Top             =   270
         Width           =   3180
      End
      Begin VB.Label Label8 
         Caption         =   "Observaciones   :"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Precio Cotizado  :"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Lead-Time Arprov.:"
         Height          =   195
         Left            =   3180
         TabIndex        =   19
         Top             =   1350
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Precio ($) :"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor :"
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   330
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Item Prov:"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F. Equivalencia :"
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "U.M. Proveedor :"
         Height          =   195
         Left            =   3180
         TabIndex        =   9
         Top             =   990
         Width           =   1215
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2520
      TabIndex        =   16
      Top             =   5700
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmManItemProv.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmManItemProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs_Lista As ADODB.Recordset
Dim StrSQL As String
Dim sTipo As String
Public varCod_item As String, varCod_Proveedor As String
Public Codigo As String, Descripcion As String
Public sUniMedDefault As String

Private Sub cmdFirst_Click()
    If Not Rs_Lista.BOF Then
        Rs_Lista.MoveFirst
    End If
End Sub

Private Sub cmdLast_Click()
    If Not Rs_Lista.EOF Then
        Rs_Lista.MoveLast
    End If
End Sub

Private Sub cmdNext_Click()
    If Not Rs_Lista.EOF Then
        Rs_Lista.MoveNext
        If Rs_Lista.EOF Then
            Rs_Lista.MoveLast
        End If
    End If
End Sub

Private Sub cmdPrevious_Click()
    If Not Rs_Lista.BOF Then
        Rs_Lista.MovePrevious
        If Rs_Lista.BOF Then
            Rs_Lista.MoveFirst
        End If
    End If
End Sub

Sub LIMPIAR_DATOS()

    txtCod_Proveedor.Text = ""
    txtDes_Proveedor.Text = ""
    txtCod_ItemProv.Text = ""
    txtFac_EquiProv.Text = "1.00"
    txtCod_UniMedProv = ""
    txtPrecio.Text = "0.00"
    txtLeadTime.Text = "0"
    txtPrecioCotizado.Text = "0.00"
    txtObservacioes.Text = ""
    
End Sub


Function VALIDA_DATOS() As Boolean
    Dim NombreTabla As String
    Dim CodigoTabla As String
    

    VALIDA_DATOS = True
    If sTipo <> "D" Then

        If sTipo = "I" Then
        
            StrSQL = "SELECT COUNT(*) FROM LG_ITEMPROV WHERE Cod_Item='" & varCod_item & "' AND Cod_Proveedor='" & Trim(txtCod_Proveedor.Text) & "' AND Cod_ItemProv='" & Trim(txtCod_ItemProv.Text) & "'"
            
            If DevuelveCampo(StrSQL, cCONNECT) <> "0" Then
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
        
        StrSQL = "SELECT count(*) FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & txtCod_Proveedor.Text & "'"
        If DevuelveCampo(StrSQL, cCONNECT) = "0" Then
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

        If Val(txtFac_EquiProv.Text) <= 0 Then
            MsgBox "El factor de equivalencia debe ser mayor a cero. Sirvase verificar", vbInformation, "Item Proveedor"
            txtFac_EquiProv.Text = ""
            txtFac_EquiProv.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If

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

    If Not Rs_Lista.EOF And Not Rs_Lista.BOF Then
        
        txtCod_Proveedor.Text = Rs_Lista("Cod_Proveedor").Value
        txtDes_Proveedor.Text = Rs_Lista("Des_Proveedor").Value
        txtCod_ItemProv.Text = Rs_Lista("Cod_ItemProv").Value
        txtFac_EquiProv.Text = Format(Rs_Lista("Fac_EquiProv").Value, "#####0.000000")
        txtCod_UniMedProv.Text = Rs_Lista("Cod_UniMedProv").Value
        txtPrecio.Text = Rs_Lista("Precio").Value
        txtPrecioCotizado.Text = Rs_Lista("Pre_Cotizado").Value
        txtObservacioes.Text = Rs_Lista("Observaciones").Value
       
    End If
End Sub

Sub HABILITA_DATOS()
    
    txtFac_EquiProv.Enabled = True
    txtCod_UniMedProv.Enabled = True
    txtPrecio.Enabled = True
    txtLeadTime.Enabled = True
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
        txtFac_EquiProv.SetFocus
    End If
    
End Sub

Sub INHABILITA_DATOS()
    
    txtCod_Proveedor.Enabled = False
    txtDes_Proveedor.Enabled = False
    txtCod_ItemProv.Enabled = False
    txtFac_EquiProv.Enabled = False
    txtCod_UniMedProv.Enabled = False
    txtPrecio.Enabled = False
    txtLeadTime.Enabled = False
    txtPrecioCotizado.Enabled = False
    txtObservacioes.Enabled = False
    
    
End Sub

Sub CARGA_GRID()
    Dim xRow As Variant
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = cCONNECT
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    'Esta cadena es para devolver el Codigo de Cliente
    StrSQL = "EXEC UP_SEL_ITEMPROV '" & varCod_item & "'" '& varCod_Proveedor & "'"
    
    xRow = DGridLista.Row
    
    Rs_Lista.Open StrSQL
    Set DGridLista.DataSource = Rs_Lista
    DGridLista.Refresh
    
    If xRow > 0 And xRow <= Rs_Lista.RecordCount Then
        DGridLista.Row = xRow
    End If

    If Rs_Lista.RecordCount > 0 Then
        DGridLista.Enabled = True
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Call Carga_Datos
    Else
        DGridLista.Enabled = False
        HabilitaMant Me.MantFunc1, "ADICIONAR"
        Call LIMPIAR_DATOS
    End If
End Sub

Function SALVAR_DATOS() As Boolean
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim StrSQL As String
    
    Con.ConnectionString = cCONNECT
    Con.Open
    
        Con.BeginTrans

        StrSQL = "EXEC UP_MAN_ITEMPROV '" & _
        sTipo & "','" & _
        varCod_item & "','" & _
        Trim(txtCod_Proveedor.Text) & "','" & _
        Trim(txtCod_ItemProv.Text) & "'," & _
        Trim(txtFac_EquiProv.Text) & ",'" & _
        Trim(txtCod_UniMedProv.Text) & "'," & _
        txtPrecio.Text & "," & _
        txtLeadTime.Text & "," & _
        txtPrecioCotizado.Text & ",'" & _
        txtObservacioes.Text & "'"

        
        
        Con.Execute StrSQL

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
       
        StrSQL = "EXEC UP_MAN_ITEMPROV '" & _
        sTipo & "','" & _
        varCod_item & "','" & _
        Trim(txtCod_Proveedor.Text) & "','" & _
        Trim(txtCod_ItemProv.Text) & "'," & _
        Trim(txtFac_EquiProv.Text) & ",'" & _
        Trim(txtCod_UniMedProv.Text) & "',0,0"
        
        Con.Execute StrSQL
    
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
                
                StrSQL = "SELECT Des_Proveedor FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & txtCod_Proveedor.Text & "'"
                txtDes_Proveedor.Text = Trim(DevuelveCampo(StrSQL, cCONNECT))
                'Strsql = "SELECT Cod_Proveedor FROM LG_PROVEEDOR WHERE Des_Proveedor = '" & txtDes_Proveedor.Text & "'"
                'txtCod_Proveedor.Text = Trim(DevuelveCampo(Strsql, cCONNECT))
                txtCod_ItemProv.SetFocus

                
        Case 2:
                Dim oTipo As New frmBusqGeneral
                Dim Rs As New ADODB.Recordset
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
                Set Rs = Nothing
                
    End Select
End Sub

Private Sub Form_Load()
    Call FormateaGrid(DGridLista)
    Call INHABILITA_DATOS
    Me.MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
    
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
            DGridLista.Enabled = False
        Case "MODIFICAR"
        
            sTipo = "U"
            HABILITA_DATOS
            'txtCod_Proveedor.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
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
                If SALVAR_DATOS Then
                    Call CARGA_GRID
                    Call INHABILITA_DATOS
                    HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                    DGridLista.Enabled = True
                    If sTipo = "I" Then
                        MantFunc1_ActionClick 0, 0, "ADICIONAR"
                    End If
                    
                End If
            End If
        Case "DESHACER"
            Call LIMPIAR_DATOS
            Call Carga_Datos
            Call INHABILITA_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
            sTipo = ""
         Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub txtCod_ItemProv_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtFac_EquiProv.SetFocus
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
        txtPrecio.SetFocus
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

Private Sub txtFac_EquiProv_LostFocus()
    If Trim(txtFac_EquiProv.Text) = "" Then
        txtFac_EquiProv.Text = "0.00"
    Else
        txtFac_EquiProv.Text = Format(txtFac_EquiProv.Text, "#######0.000000")
    End If
End Sub

Private Sub txtLeadTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtPrecioCotizado.SetFocus
    End If
End Sub

Private Sub txtLeadTime_LostFocus()
    If Trim(txtLeadTime.Text) = "" Then
        txtLeadTime.Text = "0"
    End If
End Sub

Private Sub txtObservacioes_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        MantFunc1.SetFocus
    End If
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtLeadTime.SetFocus
    End If
End Sub

Private Sub txtPrecio_LostFocus()
    If Trim(txtPrecio.Text) = "" Then
        txtPrecio.Text = "0.00"
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
