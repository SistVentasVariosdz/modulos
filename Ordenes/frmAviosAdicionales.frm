VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAviosAdicionales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Avios Adicionales"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   Icon            =   "frmAviosAdicionales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   1980
      TabIndex        =   16
      Top             =   6705
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmAviosAdicionales.frx":08CA
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalle"
      Height          =   2535
      Left            =   105
      TabIndex        =   15
      Top             =   3945
      Width           =   7395
      Begin VB.TextBox txtNro_No_Conformidad 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1155
         TabIndex        =   25
         Top             =   2010
         Width           =   1305
      End
      Begin VB.TextBox txtCod_Present 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1155
         TabIndex        =   23
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtDes_Present 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1920
         TabIndex        =   22
         Top             =   1635
         Width           =   2775
      End
      Begin VB.OptionButton optServiciosAdicionales 
         Caption         =   "Servicios Adicionales en Prenda"
         Height          =   195
         Left            =   1455
         TabIndex        =   21
         Top             =   300
         Width           =   2580
      End
      Begin VB.OptionButton optAvios 
         Caption         =   "Avios"
         Height          =   195
         Left            =   195
         TabIndex        =   20
         Top             =   300
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.TextBox txtDes_Comb 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1920
         TabIndex        =   4
         Top             =   930
         Width           =   2775
      End
      Begin VB.TextBox txtCod_Comb 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1155
         TabIndex        =   3
         Top             =   930
         Width           =   750
      End
      Begin VB.TextBox TxtCantidad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1155
         TabIndex        =   5
         Text            =   "0"
         Top             =   1290
         Width           =   765
      End
      Begin VB.TextBox txtDes_ITem 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2520
         TabIndex        =   2
         Top             =   615
         Width           =   4755
      End
      Begin VB.TextBox txtCod_Item 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1155
         TabIndex        =   1
         Top             =   615
         Width           =   1365
      End
      Begin VB.Label Label9 
         Caption         =   "Nro NO Conformidad:"
         Height          =   375
         Left            =   135
         TabIndex        =   26
         Top             =   1980
         Width           =   1005
      End
      Begin VB.Label Label8 
         Caption         =   "Presentación:"
         Height          =   225
         Left            =   135
         TabIndex        =   24
         Top             =   1710
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Combinacion:"
         Height          =   225
         Left            =   135
         TabIndex        =   19
         Top             =   1035
         Width           =   1005
      End
      Begin VB.Label Label6 
         Caption         =   "Cantidad :"
         Height          =   225
         Left            =   165
         TabIndex        =   18
         Top             =   1395
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "Item :"
         Height          =   165
         Left            =   165
         TabIndex        =   17
         Top             =   675
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3930
      Left            =   90
      TabIndex        =   6
      Top             =   30
      Width           =   7365
      Begin GridEX20.GridEX gexLista 
         Height          =   2835
         Left            =   90
         TabIndex        =   0
         Top             =   975
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   5001
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAviosAdicionales.frx":0A2A
         Column(2)       =   "frmAviosAdicionales.frx":0AF2
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAviosAdicionales.frx":0B96
         FormatStyle(2)  =   "frmAviosAdicionales.frx":0CCE
         FormatStyle(3)  =   "frmAviosAdicionales.frx":0D7E
         FormatStyle(4)  =   "frmAviosAdicionales.frx":0E32
         FormatStyle(5)  =   "frmAviosAdicionales.frx":0F0A
         FormatStyle(6)  =   "frmAviosAdicionales.frx":0FC2
         ImageCount      =   0
         PrinterProperties=   "frmAviosAdicionales.frx":10A2
      End
      Begin VB.Label Label1 
         Caption         =   "P.O.:"
         Height          =   225
         Left            =   165
         TabIndex        =   14
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Est.Cliente :"
         Height          =   225
         Left            =   150
         TabIndex        =   13
         Top             =   660
         Width           =   870
      End
      Begin VB.Label Label3 
         Caption         =   "Est.Propio :"
         Height          =   240
         Left            =   3990
         TabIndex        =   12
         Top             =   645
         Width           =   885
      End
      Begin VB.Label Label4 
         Caption         =   "O.P.:"
         Height          =   225
         Left            =   4005
         TabIndex        =   11
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblPO 
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
         Left            =   1155
         TabIndex        =   10
         Top             =   315
         Width           =   2445
      End
      Begin VB.Label lblOP 
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
         Left            =   4920
         TabIndex        =   9
         Top             =   270
         Width           =   2250
      End
      Begin VB.Label lblEstCli 
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
         Left            =   1140
         TabIndex        =   8
         Top             =   645
         Width           =   2445
      End
      Begin VB.Label lblEstPro 
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
         Left            =   4920
         TabIndex        =   7
         Top             =   645
         Width           =   2250
      End
   End
End
Attribute VB_Name = "frmAviosAdicionales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Cliente As String
Public sCod_PurOrd As String
Public sCod_LotPurOrd As String
Public sCod_EstCli As String
Public sCod_EstPro As String

Public Codigo As String
Public Descripcion As String

Dim StrSql As String
Dim stipo As String

Public Function CARGA_GRID() As Boolean
On Error GoTo errores
Dim vBookmark As Variant

vBookmark = gexLista.Row
gexLista.ClearFields

StrSql = "EXEC SM_SEL_LOTESTPRO_AVIOS_ADICIONALES '" & sCod_Cliente & "','" & _
                                                        sCod_PurOrd & "','" & _
                                                        sCod_LotPurOrd & "','" & _
                                                        sCod_EstCli & "','" & _
                                                        sCod_EstPro & "'"


Set gexLista.ADORecordset = CargarRecordSetDesconectado(StrSql, cCONNECT)

gexLista.Row = vBookmark

gexLista.Columns("codigo").Width = 1000
gexLista.Columns("descripcion").Width = 3500

Exit Function

errores:
    errores Err.Number
End Function


Private Sub Form_Load()
    Me.MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Private Sub gexLista_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    txtCod_Item.Text = gexLista.value(gexLista.Columns("codigo").Index)
    txtDes_ITem.Text = gexLista.value(gexLista.Columns("descripcion").Index)
    txtCod_Comb.Text = gexLista.value(gexLista.Columns("cod.comb.").Index)
    txtDes_Comb.Text = gexLista.value(gexLista.Columns("des.comb.").Index)
    TxtCantidad.Text = gexLista.value(gexLista.Columns("cantidad").Index)
    If gexLista.value(gexLista.Columns("tipo").Index) = "I" Then
        optAvios.value = True
        optServiciosAdicionales.value = False
    Else
        optAvios.value = False
        optServiciosAdicionales.value = True
    End If
    txtCod_Present.Text = gexLista.value(gexLista.Columns("cod_present").Index)
    txtDes_Present.Text = gexLista.value(gexLista.Columns("des_present").Index)
    txtNro_No_Conformidad.Text = gexLista.value(gexLista.Columns("Nro_No_Conformidad").Index)
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim eliminar As Integer
    Select Case ActionName
        Case "ADICIONAR"
            stipo = "I"
            Call LIMPIA_DATOS
            Call HABILITA_DATOS(True)
            txtCod_Item.Enabled = True
            txtDes_ITem.Enabled = True
            txtCod_Comb.Enabled = True
            txtDes_Comb.Enabled = True
            txtCod_Present.Enabled = True
            txtDes_Present.Enabled = True
            txtNro_No_Conformidad.Enabled = True
            Me.txtCod_Item.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Case "MODIFICAR"
            stipo = "U"
            Call HABILITA_DATOS(True)
            Me.TxtCantidad.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Case "ELIMINAR"
            eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Combinación-Detalle")
            If eliminar = vbYes Then
                stipo = "D"
                If VALIDA_DATOS Then
                    Call SALVAR_DATOS
                    Call CARGA_GRID
                    stipo = ""
                End If
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                Call SALVAR_DATOS
                Call CARGA_GRID
                Call HABILITA_DATOS(False)
                txtCod_Item.Enabled = False
                txtDes_ITem.Enabled = False
                txtCod_Comb.Enabled = False
                txtDes_Comb.Enabled = False
                txtCod_Present.Enabled = False
                txtDes_Present.Enabled = False
                txtNro_No_Conformidad.Enabled = False
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                stipo = ""
            End If
        Case "DESHACER"
            Call LIMPIA_DATOS
            Call CARGA_GRID
            Call HABILITA_DATOS(False)
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            stipo = ""
         Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub TxtCantidad_GotFocus()
    SelectionText TxtCantidad
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If optAvios Then
            Me.MantFunc1.SetFocus
        Else
            If txtCod_Present.Enabled Then
                Me.txtCod_Present.SetFocus
            Else
                If txtNro_No_Conformidad.Enabled Then
                    Me.txtNro_No_Conformidad.SetFocus
                End If
            End If
        End If
    Else
        Call SoloNumeros(TxtCantidad, KeyAscii, False)
    End If
End Sub

Private Sub txtCod_Comb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Comb.Text) = "" Then
            BUSCA_COMBINACION (3)
        Else
            BUSCA_COMBINACION (1)
        End If
    End If
End Sub

Private Sub txtcod_item_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtCod_Item.Text)) < 3 Then
            If Trim(txtCod_Item.Text) = "" Then
                txtDes_ITem.Text = ""
                txtDes_ITem.SetFocus
            Else
                BUSCA_ITEM (3)
            End If
        Else
            BUSCA_ITEM (1)
        End If
    End If
End Sub

Private Sub txtCod_Present_GotFocus()
    SelectionText txtCod_Present
End Sub

Private Sub txtCod_Present_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(txtCod_Present.Text) = "" Then
            BUSCA_PRESENTACION (3)
        Else
            BUSCA_PRESENTACION (1)
        End If
        txtNro_No_Conformidad.SetFocus
    End If
End Sub

Private Sub txtDes_Comb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDes_Comb.Text) = "" Then
            BUSCA_COMBINACION (3)
        Else
            BUSCA_COMBINACION (2)
        End If
    End If
End Sub

Private Sub txtdes_item_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDes_ITem.Text) = "" Then
            TxtCantidad.SetFocus
        Else
            BUSCA_ITEM (2)
        End If
    End If
End Sub

Sub LIMPIA_DATOS()
    optAvios.value = True
    txtCod_Item.Text = ""
    txtDes_ITem.Text = ""
    txtCod_Comb.Text = ""
    txtDes_Comb.Text = ""
    TxtCantidad.Text = 0
    txtCod_Present.Text = 0
    txtDes_Present.Text = ""
End Sub

Private Sub TxtCantidad_LostFocus()
    If Trim(TxtCantidad.Text) = "" Then TxtCantidad.Text = 0
End Sub

Sub HABILITA_DATOS(sEstado As Boolean)
    TxtCantidad.Enabled = sEstado
End Sub

Sub SALVAR_DATOS()
On Error GoTo errores

Dim sCod_CompEst  As String

If optAvios Then
    txtCod_Present.Text = 0
    sCod_CompEst = "GEN"
Else
    sCod_CompEst = "SAD"
End If

StrSql = "EXEC UP_MAN_TG_LOTESTPRO_AVIOS '" & stipo & "','" & _
                                                sCod_Cliente & "','" & _
                                                sCod_PurOrd & "','" & _
                                                sCod_LotPurOrd & "','" & _
                                                sCod_EstCli & "','" & _
                                                sCod_EstPro & "','" & _
                                                txtCod_Item.Text & "','" & _
                                                txtCod_Comb.Text & "'," & _
                                                TxtCantidad.Text & ",'" & sCod_CompEst & "'," & txtCod_Present & ",'" & txtNro_No_Conformidad.Text & "'"
                                                
Call ExecuteCommandSQL(cCONNECT, StrSql)

Exit Sub
errores:
    errores Err.Number
    'ErrorHandler Err, "SALVAR_DATOS"
End Sub

Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    
    If Trim(txtCod_Item.Text) = "" Then
        MsgBox "Ingrese el Item", vbInformation, "Aviso"
        txtCod_Item.SetFocus
        VALIDA_DATOS = False
        Exit Function
    End If

    If CDbl(TxtCantidad.Text) < 1 Then
        MsgBox "Ingrese una cantidad valida", vbInformation, "Aviso"
        TxtCantidad.SetFocus
        VALIDA_DATOS = False
        Exit Function
    End If

    If optServiciosAdicionales.value And RTrim(txtCod_Present.Text) = 0 Then
        MsgBox "Ingrese una presentación válida", vbInformation, "Aviso"
        If txtCod_Present.Enabled Then
            txtCod_Present.SetFocus
        End If
        VALIDA_DATOS = False
        Exit Function
    End If
    
End Function

Public Sub BUSCA_ITEM(Tipo As Integer)
    Select Case Tipo
        Case 1:
                    StrSql = Trim(DevuelveCampo("Select dbo.uf_devuelvecodigo(6," & IIf(Trim(txtCod_Item.Text) = "", 0, Mid(txtCod_Item, 3)) & ")", cCONNECT))
                    txtCod_Item.Text = UCase(Left(txtCod_Item, 2) & StrSql)
                    StrSql = "SELECT Des_Item FROM LG_ITEM WHERE Cod_Item = '" & Trim(Me.txtCod_Item.Text) & "'"
                    Me.txtDes_ITem.Text = Trim(DevuelveCampo(StrSql, cCONNECT))
                    Me.txtCod_Comb.SetFocus
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "SELECT Cod_Item as Codigo,Des_Item as Descripcion FROM LG_ITEM WHERE Des_Item like '%" & Trim(Me.txtDes_ITem.Text) & "%' order by cod_item"
                    Else
                        oTipo.SQuery = "SELECT Cod_Item as Codigo,Des_Item as Descripcion FROM LG_ITEM WHERE cod_item like '%" & txtCod_Item.Text & "%' order by cod_item"
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.Show 1
                    If Codigo <> "" Then
                         Me.txtCod_Item.Text = Trim(Codigo)
                         Me.txtDes_ITem.Text = Trim(Descripcion)
                         Codigo = "": Descripcion = ""
                         Me.txtCod_Comb.SetFocus
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    
End Sub

Public Sub BUSCA_COMBINACION(Tipo As Integer)
    Select Case Tipo
        Case 1:
                    StrSql = "SELECT Des_Comb FROM lg_itemcomb WHERE Cod_Comb = '" & Trim(Me.txtCod_Comb.Text) & "'"
                    Me.txtDes_Comb.Text = Trim(DevuelveCampo(StrSql, cCONNECT))
                    Me.TxtCantidad.SetFocus
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "SELECT Cod_Comb as Codigo,Des_Comb as Descripcion FROM lg_itemcomb WHERE cod_item='" & txtCod_Item.Text & "' and  Des_Comb like '%" & Trim(Me.txtDes_Comb.Text) & "%' order by cod_comb"
                    Else
                        oTipo.SQuery = "SELECT Cod_Comb as Codigo,Des_Comb as Descripcion FROM lg_itemcomb WHERE cod_item='" & txtCod_Item.Text & "' order by cod_comb"
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.Show 1
                    If Codigo <> "" Then
                         Me.txtCod_Comb.Text = Trim(Codigo)
                         Me.txtDes_Comb.Text = Trim(Descripcion)
                         Codigo = "": Descripcion = ""
                         Me.TxtCantidad.SetFocus
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    
End Sub

Private Sub txtNro_No_Conformidad_GotFocus()
    SelectionText txtNro_No_Conformidad
End Sub

Private Sub txtNro_No_Conformidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        MantFunc1.SetFocus
    End If
End Sub


Public Sub BUSCA_PRESENTACION(Tipo As Integer)
    Select Case Tipo
        Case 1:
                    StrSql = "SELECT Des_Present  FROM es_estpropre WHERE cod_estpro = '" & sCod_EstPro & "' and Cod_present = '" & Trim(Me.txtCod_Present.Text) & "'"
                    Me.txtDes_Present.Text = Trim(DevuelveCampo(StrSql, cCONNECT))
                    Me.MantFunc1.SetFocus
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "SELECT Cod_Present as Codigo,Des_present as Descripcion FROM es_estpropre WHERE cod_estpro ='" & sCod_EstPro & "' and  Des_present like '%" & Trim(Me.txtDes_Present.Text) & "%' order by cod_present"
                    Else
                        oTipo.SQuery = "SELECT Cod_Present as Codigo,Des_present as Descripcion FROM es_estpropre WHERE cod_estpro ='" & sCod_EstPro & "' order by cod_present "
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.Show 1
                    If Codigo <> "" Then
                         Me.txtCod_Present.Text = Trim(Codigo)
                         Me.txtDes_Present.Text = Trim(Descripcion)
                         Codigo = "": Descripcion = ""
                         Me.MantFunc1.SetFocus
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    
End Sub

