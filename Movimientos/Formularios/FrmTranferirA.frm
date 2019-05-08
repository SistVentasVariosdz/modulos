VERSION 5.00
Begin VB.Form FrmTranferirA 
   Caption         =   "Transferir"
   ClientHeight    =   2415
   ClientLeft      =   2910
   ClientTop       =   6435
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   4935
      TabIndex        =   16
      Top             =   1890
      Width           =   1170
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   3570
      TabIndex        =   15
      Top             =   1890
      Width           =   1170
   End
   Begin VB.Frame Fradetalle 
      Caption         =   "Datos a Transferir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   0
      TabIndex        =   0
      Tag             =   "Detail"
      Top             =   0
      Width           =   9510
      Begin VB.TextBox TxtItem 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1380
         MaxLength       =   8
         TabIndex        =   19
         Top             =   315
         Width           =   945
      End
      Begin VB.TextBox TxtDesitem 
         Height          =   315
         Left            =   2340
         TabIndex        =   18
         Top             =   315
         Width           =   1665
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   315
         Left            =   4020
         TabIndex        =   17
         Top             =   315
         Width           =   345
      End
      Begin VB.ComboBox CmbTalla 
         Height          =   315
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   630
         Width           =   2715
      End
      Begin VB.ComboBox CmbCombinacion 
         Height          =   315
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   300
         Width           =   2715
      End
      Begin VB.ComboBox CmbEstilo 
         Height          =   315
         ItemData        =   "FrmTranferirA.frx":0000
         Left            =   6720
         List            =   "FrmTranferirA.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   2715
      End
      Begin VB.ComboBox CmbDestino 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   2355
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   315
         Left            =   4020
         TabIndex        =   4
         Top             =   630
         Width           =   345
      End
      Begin VB.TextBox TxtDetalle 
         Height          =   315
         Left            =   2340
         TabIndex        =   3
         Top             =   630
         Width           =   1665
      End
      Begin VB.TextBox CmbColor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1380
         MaxLength       =   7
         TabIndex        =   2
         Top             =   630
         Width           =   945
      End
      Begin VB.TextBox TxtCodProv 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6720
         TabIndex        =   1
         Top             =   1260
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   20
         Top             =   420
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Medida:"
         Height          =   195
         Index           =   4
         Left            =   5610
         TabIndex        =   14
         Top             =   705
         Width           =   570
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   270
         TabIndex        =   13
         Tag             =   "Hilado :"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estilo:"
         Height          =   195
         Index           =   3
         Left            =   5640
         TabIndex        =   12
         Top             =   1035
         Width           =   420
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   270
         TabIndex        =   11
         Tag             =   "Hilado :"
         Top             =   1035
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Combinacion:"
         Height          =   195
         Index           =   0
         Left            =   5565
         TabIndex        =   10
         Top             =   405
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cod Prov.:"
         Height          =   195
         Index           =   2
         Left            =   5640
         TabIndex        =   9
         Top             =   1335
         Width           =   750
      End
   End
End
Attribute VB_Name = "FrmTranferirA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Codigo
Public Descripcion
Public Paso As Boolean
Public item As String
Public comb As String
Public Color As String
Public Estilo As String
Public medida As String
Public Destino As String
Public cod_prov As String

Public itemAnt As String
Public combAnt As String
Public colorAnt As String
Public estiloAnt As String
Public medidaAnt As String
Public destinoAnt As String
Public cod_provAnt As String

Public var_tipo As String

Public Cod_Almacen As String
Public Num_MovStk As String
Public Ser_OrdComp As String
Public Cod_OrdComp As String
Public cod_tipmov As String

Public varTallaProv As String

Dim Rs As New Recordset

Sub LlenarCombos()
LlenaCombo CmbDestino, "select des_destino+space(100)+cod_destino from tg_destino order by 1", cConnect
LlenaCombo CmbEstilo, "select rtrim(cod_estcli)+'  -  '+des_estcli+space(100)+cod_estcli from tg_estcli order by 1", cConnect
End Sub

Private Sub CmbColor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CmbColor <> "" Then
        'CmbColor = DevuelveCampo("select dbo.uf_devuelvecodigo(6," & CmbColor & ")", cConnect)
        If ExisteCampo("Cod_color", "lb_color", CmbColor, cConnect, True) Then
            TxtDetalle = DevuelveCampo("Select Des_color from lb_color where Cod_color='" & CmbColor & "'", cConnect)
        Else
            MsgBox "El codigo no existe", vbInformation
        End If
    Else
         Command1_Click
    End If
End If
End Sub

Public Sub CmbCombinacion_DropDown()
LlenaCombo Me.CmbCombinacion, "select Des_Comb+space(100)+Cod_Comb from lg_itemcomb where cod_item='" & Me.TxtItem & "'", cConnect
End Sub

Public Sub CmbTalla_DropDown()
Dim Tot As Integer
Tot = DevuelveCampo("Select count(*) from lg_itemmed where cod_item='" & TxtItem & "'", cConnect)

If Tot > 0 Then
    LlenaCombo CmbTalla, "Select Descripcion+space(100)+Cod_Medida from lg_itemmed where cod_item='" & TxtItem & "'  order by 1", cConnect
Else
    LlenaCombo CmbTalla, "select cod_talla from tg_talla  order by 1", cConnect
End If

End Sub

Private Sub cmdAceptar_Click()
If Valida_Diferente = False Then
    MsgBox "Imposible transferir, mismos valores", vbCritical
    Exit Sub
Else
    If var_tipo = "E" Then
         If TxtCodProv <> "" Then
             If Not ExisteCampo("cod_prov", "lg_movistkitem", TxtCodProv, cConnect) Then
                 MsgBox "Cod. Prov. no valido", vbInformation
                 Exit Sub
             Else
                 If VALIDA_PROV = False Then
                     MsgBox "Cod. Prov. no válido", vbInformation
                     Exit Sub
                 End If
             End If
         End If
     ElseIf var_tipo = "S" Then
         If TxtCodProv <> "" Then
'             If Not ExisteCampo("cod_prov", "LG_STOCKSITEM_PROV", TxtCodProv, cConnect) Then
'                 MsgBox "Cod. Prov. no valido", vbInformation
'                 Exit Sub
'             Else
'                 If VALIDA_PROV_SALIDA = False Then
'                     MsgBox "Cod. Prov. no válido", vbInformation
'                     Exit Sub
'                 End If
'             End If
'         Else
'             MsgBox "Cod. Prov. ES OBLIGATORIO", vbInformation
         End If
     End If
    
    
    Unload Me
End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Set frmBusqGeneral.oParent = Me
Codigo = ""
Descripcion = ""
frmBusqGeneral.sQuery = "select Cod_color as Codigo ,Des_color as Nombre from lb_color order by 2"
frmBusqGeneral.CARGAR_DATOS
frmBusqGeneral.Show 1

CmbColor = Codigo
TxtDetalle = Descripcion
End Sub

Private Sub Command2_Click()
Set frmBusqGeneral.oParent = Me

frmBusqGeneral.sQuery = "select Cod_item as Codigo ,Des_item as Nombre from lg_item where Des_item<>'' order by 2"
frmBusqGeneral.CARGAR_DATOS
frmBusqGeneral.Show 1

TxtItem = Codigo
TxtDesitem = Descripcion

End Sub

Private Sub Form_Load()
'LlenarCombos
End Sub

Private Sub TxtCodProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If var_tipo = "E" Then
        MUESTRA_AYUDA
    ElseIf var_tipo = "S" Then
        MUESTRA_AYUDA_SALIDA
    End If
End If
End Sub

Private Sub TxtItem_KeyPress(KeyAscii As Integer)
On Error GoTo hand
Dim Temp As String
Dim strSQL As String

If KeyAscii = 13 Then
    If Len(TxtItem.Text) < 3 Then
         Set frmBusqGeneral.oParent = Me
         Codigo = ""
         Descripcion = ""
         frmBusqGeneral.sQuery = "select Cod_Item AS Codigo,des_item as Descripcion from lg_item where Cod_item like '" & TxtItem & "%'"
         frmBusqGeneral.CARGAR_DATOS
         frmBusqGeneral.Show 1
         TxtDesitem = Descripcion
         TxtItem = Codigo
         Temp = TxtItem
         
         If Codigo <> "" Then
            GoTo otro
         Else
            Exit Sub
         End If
    End If

    If Len(TxtItem.Text) > 2 Then Temp = Trim(DevuelveCampo("Select dbo.uf_devuelvecodigo(6," & IIf(Trim(TxtItem) = "", 0, Mid(TxtItem, 3)) & ")", cConnect))

    Temp = Left(TxtItem, 2) & Temp
        If DevuelveCampo("select count(*) from lg_item where cod_item ='" & Temp & "'", cConnect) > 0 Then
            Me.TxtDesitem = DevuelveCampo("select Des_Item from lg_item where cod_item ='" & Temp & "'", cConnect)
            TxtItem = Temp
        Else
            MsgBox "Codigo no existe", vbInformation
            Me.TxtDesitem = ""
            Exit Sub
        End If
        
otro:
    'Strsql = "SELECT COUNT(*) FROM LG_TIPOSMOV WHERE TIP_ITEM = 'I' AND COD_CLAMOV='S' AND ISNULL(Flg_SecOrd,'') <> '*' AND RTRIM(Cod_ClaOrdComp) = '' AND Cod_TipMov = '" & Me.Cod_TipMov & "'"
    strSQL = "EXEC UP_VERIFICA_MOV_AVIO_SAL '" & Me.cod_tipmov & "'"
    If Val(DevuelveCampo(strSQL, cConnect)) = 1 Then

        varTallaProv = ""

        Load frmListaStocksAvios
        frmListaStocksAvios.varCOD_ALMACEN = Me.Cod_Almacen
        frmListaStocksAvios.varCod_Item = Me.TxtItem.Text
        frmListaStocksAvios.Caption = "Stocks del : " & Me.TxtItem & " - " & Me.TxtDesitem
        frmListaStocksAvios.varOpcionBusq = "2"
        Set frmListaStocksAvios.oParent = Me
        frmListaStocksAvios.CARGA_GRID
        frmListaStocksAvios.Show 1

        Set frmListaStocksAvios = Nothing
        
    End If
End If
Exit Sub
hand:

End Sub

Sub MUESTRA_AYUDA_SALIDA()
Set frmBusqGeneral2.oParent = Me
        Codigo = ""
        Descripcion = ""
         frmBusqGeneral2.sQuery = "MUESTRA_AYUDA_PROV_SALIDAS '" & Cod_Almacen & "','" & _
                            Me.TxtItem & "','" & Me.CmbCombinacion & "','" & Me.CmbColor & "','" & Me.CmbTalla & "','" & _
                            Me.CmbEstilo & "','" & Me.CmbDestino & "'"
                                    
         'frmBusqGeneral2.sQuery = "select Cod_Prov from lg_ordcompitem where COD_ITEM = '" & TxtItem & "' AND COD_COLOR='" & CmbColor & "' AND COD_COMB='" & CmbCombinacion & "' AND COD_TALLA='" & CmbTalla & "' AND COD_DESTINO='" & CmbDestino & "' AND COD_ESTCLI='" & CmbEstilo & "' and cod_prov<>''"
         frmBusqGeneral2.CARGAR_DATOS
         frmBusqGeneral2.Show 1
         If Codigo <> "" Then
            TxtCodProv = Trim(Codigo)
        End If
End Sub

Sub MUESTRA_AYUDA()
Set frmBusqGeneral2.oParent = Me
        Codigo = ""
        Descripcion = ""
         frmBusqGeneral2.sQuery = "MUESTRA_AYUDA_PROV_ENTRADA '" & Me.Cod_Almacen & "','" & _
                                    Me.Num_MovStk & "','" & Me.Ser_OrdComp & "','" & Me.Cod_OrdComp & "','" & _
                                    Me.TxtItem & "','" & Me.CmbCombinacion & "','" & Me.CmbColor & "','" & Me.CmbTalla & "','" & _
                                    Me.CmbEstilo & "','" & Me.CmbDestino & "'"
                                    
         'frmBusqGeneral2.sQuery = "select Cod_Prov from lg_ordcompitem where COD_ITEM = '" & TxtItem & "' AND COD_COLOR='" & CmbColor & "' AND COD_COMB='" & CmbCombinacion & "' AND COD_TALLA='" & CmbTalla & "' AND COD_DESTINO='" & CmbDestino & "' AND COD_ESTCLI='" & CmbEstilo & "' and cod_prov<>''"
         frmBusqGeneral2.CARGAR_DATOS
         frmBusqGeneral2.Show 1
         If Codigo <> "" Then
            TxtCodProv = Trim(Codigo)
        End If
End Sub

Function VALIDA_PROV() As Boolean
    Set Rs = Nothing
    Rs.CursorLocation = adUseClient
    Rs.Open "VALIDA_PROV '" & Me.Cod_Almacen & "','" & _
                            Me.Num_MovStk & "','" & Me.Ser_OrdComp & "','" & Me.Cod_OrdComp & "','" & _
                            Me.TxtItem & "','" & Trim(Right(Me.CmbCombinacion, 3)) & "','" & Trim(Me.CmbColor) & "','" & Trim(Right(Me.CmbTalla, 10)) & "','" & _
                            Trim(Right(Me.CmbEstilo, 25)) & "','" & Right(meCmbDestino.Text, 4) & "','" & TxtCodProv & "'", cConnect, 3, 3
    
    If Rs.RecordCount <= 0 Then
        'MsgBox "Cod. Prov. no valido", vbInformation
        VALIDA_PROV = False
    Else
        VALIDA_PROV = True
    End If
End Function

Function VALIDA_PROV_SALIDA() As Boolean
    Set Rs = Nothing
    Rs.CursorLocation = adUseClient
    Rs.Open "VALIDA_PROV_SALIDAS '" & Cod_Almacen & "','" & _
                            Me.TxtItem & "','" & Trim(Right(Me.CmbCombinacion, 3)) & "','" & Trim(Me.CmbColor) & "','" & Trim(Right(Me.CmbTalla, 10)) & "','" & _
                            Trim(Right(Me.CmbEstilo, 25)) & "','" & Right(Me.CmbDestino.Text, 4) & "','" & Me.TxtCodProv & "'", cConnect, 3, 3
    If Rs.RecordCount <= 0 Then
        'MsgBox "Cod. Prov. no valido", vbInformation
        VALIDA_PROV_SALIDA = False
    Else
        VALIDA_PROV_SALIDA = True
    End If
                            
End Function


Function Valida_Diferente() As Boolean
    item = TxtItem.Text
    comb = Trim(Right(Me.CmbCombinacion, 3))
    Color = Trim(Me.CmbColor.Text)
    Estilo = Trim(Right(Me.CmbEstilo, 25))
    medida = Trim(Right(Me.CmbTalla, 10))
    Destino = Right(Me.CmbDestino.Text, 4)
    cod_prov = Me.TxtCodProv.Text
    
    With FrmDetalleStock
        If itemAnt & combAnt & colorAnt & estiloAnt & medidaAnt & destinoAnt & cod_provAnt = item & comb & Color & Estilo & medida & Destino & cod_prov Then
            Valida_Diferente = False
        Else
            Valida_Diferente = True
        End If
        
    End With
End Function
