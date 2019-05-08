VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form frmTX_Rapport_Detalle 
   Caption         =   "Detalle del Rapport"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   7035
      TabIndex        =   7
      Top             =   6740
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~REPORTE~Verdadero~Verdadero~&Reporte Resumido~0~0~1~~0~Falso~Falso~&Reporte Resumido~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Fradetalle 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   10
      Tag             =   "Detail"
      Top             =   2940
      Width           =   8325
      Begin VB.TextBox TxtCod_HiladoEstructurado 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   30
         Top             =   3300
         Width           =   1515
      End
      Begin VB.ComboBox cboCod_HilTel 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1380
         Width           =   5325
      End
      Begin VB.TextBox TxtHiloAntiguo 
         Enabled         =   0   'False
         Height          =   330
         Left            =   6840
         TabIndex        =   27
         Top             =   3360
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtCod_Hiltel 
         Height          =   330
         Left            =   1680
         TabIndex        =   2
         Top             =   1365
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtSecuencia 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1680
         TabIndex        =   1
         Top             =   1000
         Width           =   1250
      End
      Begin VB.TextBox TxtDes_Comb 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   330
         Left            =   3045
         TabIndex        =   26
         Top             =   630
         Width           =   5160
      End
      Begin VB.TextBox TxtDes_Rapport 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   330
         Left            =   3045
         TabIndex        =   25
         Top             =   210
         Width           =   5160
      End
      Begin VB.TextBox TxtDes_Color 
         Height          =   330
         Left            =   3675
         TabIndex        =   24
         Top             =   1785
         Width           =   4425
      End
      Begin VB.TextBox TxtDes_Hiltel 
         Height          =   330
         Left            =   3675
         TabIndex        =   23
         Top             =   1365
         Visible         =   0   'False
         Width           =   4425
      End
      Begin VB.TextBox TxtPorcentaje 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   4935
         TabIndex        =   21
         Top             =   1000
         Width           =   1260
      End
      Begin VB.TextBox TxtNroPasadas 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   2205
         Width           =   1575
      End
      Begin VB.TextBox TxtObservacion 
         Height          =   600
         Left            =   1680
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   2610
         Width           =   6405
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   330
         Left            =   3075
         TabIndex        =   17
         Top             =   1770
         Width           =   435
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   330
         Left            =   3075
         TabIndex        =   16
         Top             =   1365
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtCod_comb 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
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
         Left            =   1680
         TabIndex        =   9
         Top             =   630
         Width           =   1250
      End
      Begin VB.TextBox txtRapport 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
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
         Left            =   1665
         MaxLength       =   8
         TabIndex        =   8
         Top             =   240
         Width           =   1250
      End
      Begin VB.TextBox txtCod_Color 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   1785
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hilado Antiguo"
         Height          =   195
         Left            =   210
         TabIndex        =   28
         Top             =   3330
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje:"
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
         Left            =   3675
         TabIndex        =   22
         Top             =   1050
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Pasadas :"
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
         Left            =   315
         TabIndex        =   20
         Tag             =   "Porcentaje :"
         Top             =   2280
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones:"
         Height          =   195
         Left            =   255
         TabIndex        =   19
         Top             =   2625
         Width           =   1110
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Comb.:"
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
         Index           =   1
         Left            =   225
         TabIndex        =   15
         Tag             =   "Mat. Prima :"
         Top             =   690
         Width           =   870
      End
      Begin VB.Label lblCod_Item 
         AutoSize        =   -1  'True
         Caption         =   "RN:"
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
         Index           =   0
         Left            =   225
         TabIndex        =   14
         Tag             =   "Hilado :"
         Top             =   315
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Secuencia :"
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
         Left            =   210
         TabIndex        =   13
         Top             =   1050
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Color:"
         Height          =   195
         Left            =   255
         TabIndex        =   12
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Hilado Nuevo:"
         Height          =   195
         Left            =   225
         TabIndex        =   11
         Top             =   1440
         Width           =   1020
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
      Height          =   2910
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8355
      Begin GridEX20.GridEX gexDetalleRapport 
         Height          =   2535
         Left            =   105
         TabIndex        =   18
         Top             =   210
         Width           =   8070
         _ExtentX        =   14235
         _ExtentY        =   4471
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         BackColorBkg    =   -2147483634
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmTX_Rapport_Detalle.frx":0000
         Column(2)       =   "frmTX_Rapport_Detalle.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmTX_Rapport_Detalle.frx":016C
         FormatStyle(2)  =   "frmTX_Rapport_Detalle.frx":02A4
         FormatStyle(3)  =   "frmTX_Rapport_Detalle.frx":0354
         FormatStyle(4)  =   "frmTX_Rapport_Detalle.frx":0408
         FormatStyle(5)  =   "frmTX_Rapport_Detalle.frx":04E0
         FormatStyle(6)  =   "frmTX_Rapport_Detalle.frx":0598
         ImageCount      =   0
         PrinterProperties=   "frmTX_Rapport_Detalle.frx":0678
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2100
      TabIndex        =   6
      Top             =   6720
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmTX_Rapport_Detalle.frx":0850
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmTX_Rapport_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Rapport As String
Public Comb As String
Public secuencia As String
Dim strSQL As String
Dim sTipo As String
Dim vMessage As String
Public Codigo As String
Public Descripcion As String
Public Tela As String

Sub CARGA_GRID()
On Error GoTo ErrCargaGrid
    strSQL = "up_sel_tx_rapport_detalle " & txtRapport & ",'" & txtCod_comb & "'"
    Set gexDetalleRapport.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
    
    gexDetalleRapport.Columns("rapport_comb").Width = "700"
    gexDetalleRapport.Columns("secuencia").Width = "600"
    gexDetalleRapport.Columns("rapport_number").Width = "500"
    
    gexDetalleRapport.Columns("hilado").Width = "900"
    gexDetalleRapport.Columns("des_hiltel").Width = "1500"
    gexDetalleRapport.Columns("porcentaje").Width = "700"
    gexDetalleRapport.Columns("cod_color").Width = "800"
    gexDetalleRapport.Columns("nro_pasadas").Width = "900"
    gexDetalleRapport.Columns("des_color").Width = "1500"
    gexDetalleRapport.Columns("observaciones").Width = "2200"
    
    gexDetalleRapport.Columns("cod_hiltel").Visible = False
    gexDetalleRapport.Columns("rapport_number").Caption = "RN"
    gexDetalleRapport.Columns("rapport_comb").Caption = "Comb."
    gexDetalleRapport.Columns("porcentaje").Caption = "Porcen."
    DESHABILITA
    Exit Sub
ErrCargaGrid:
ErrorHandler Err, "Carga_Grid"
End Sub

Private Sub cboCod_HilTel_Click()
TxtCod_HiladoEstructurado = Trim(Right(cboCod_HilTel.Text, 10))
End Sub

Private Sub Command1_Click()
Dim oTipo As New frmBusqGeneral
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT COD_HILADO_ESTRUCTURADO as Código, Des_HILTEL as Descripción FROM IT_HILADO where cod_hilado_estructurado<>''"
    
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtCod_Hiltel.Text = Codigo
        TxtDes_Hiltel.Text = Descripcion
        strSQL = "select cod_hiltel from it_hilado where cod_hilado_estructurado='" & txtCod_Hiltel.Text & "'"
        TxtHiloAntiguo.Text = DevuelveCampo(strSQL, cCONNECT)
    Else
        txtCod_Hiltel.Text = ""
        TxtDes_Hiltel.Text = ""
        TxtHiloAntiguo.Text = ""
    End If
    Codigo = ""
    Descripcion = ""
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub

Sub busca_hilo_des()
Dim oTipo As New frmBusqGeneral
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT COD_HILADO_ESTRUCTURADO as Código, Des_HILTEL as Descripción FROM IT_HILADO WHERE DES_HILTEL LIKE '%" & Trim(TxtDes_Hiltel) & "%' and cod_hilado_estructurado<>''"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtCod_Hiltel.Text = Codigo
        TxtDes_Hiltel.Text = Descripcion
        strSQL = "select cod_hiltel from it_hilado where cod_hilado_estructurado='" & txtCod_Hiltel.Text & "'"
        TxtHiloAntiguo.Text = DevuelveCampo(strSQL, cCONNECT)
    Else
        txtCod_Hiltel.Text = ""
        TxtDes_Hiltel.Text = ""
        TxtHiloAntiguo.Text = ""
    End If
    Codigo = ""
    Descripcion = ""
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub


Private Sub Command2_Click()
    Call Me.BUSCA_COLOR(3)
End Sub

Private Sub Form_Load()
DESHABILITA

strSQL = "SM_Muestra_It_Hilado"
Call LlenaCombo(cboCod_HilTel, strSQL, cCONNECT)
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    frmShowTX_Rapport_Detalle.Rapport_Number = Me.txtRapport
    frmShowTX_Rapport_Detalle.CARGA_GRID
    frmShowTX_Rapport_Detalle.Show 1
End Sub

Private Sub gexDetalleRapport_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    If gexDetalleRapport.RowCount > 0 Then
        txtRapport.Text = gexDetalleRapport.Value(gexDetalleRapport.Columns("rapport_number").Index)
        txtCod_comb.Text = Trim(gexDetalleRapport.Value(gexDetalleRapport.Columns("rapport_comb").Index))
    End If
    
    If txtRapport.Text <> "" Then
        TxtDes_Rapport = DevuelveCampo("SELECT DESCRIPCION FROM TX_RAPPORT WHERE RAPPORT_NUMBER =" & txtRapport, cCONNECT)
        TxtDes_Comb = DevuelveCampo("SELECT DESCRIPCION FROM TX_RAPPORT_comb WHERE RAPPORT_NUMBER =" & txtRapport & " AND RAPPORT_COMB='" & txtCod_comb & "'", cCONNECT)
    End If
    
    txtSecuencia.Text = Trim(gexDetalleRapport.Value(gexDetalleRapport.Columns("secuencia").Index))
    
    txtCod_Hiltel.Text = Trim(gexDetalleRapport.Value(gexDetalleRapport.Columns("Hilado").Index))
    TxtDes_Hiltel.Text = Trim(gexDetalleRapport.Value(gexDetalleRapport.Columns("des_hiltel").Index))
    
    TxtHiloAntiguo.Text = Trim(gexDetalleRapport.Value(gexDetalleRapport.Columns("cod_hiltel").Index))
    
    txtCod_Color.Text = Trim(gexDetalleRapport.Value(gexDetalleRapport.Columns("cod_color").Index))
    TxtDes_Color.Text = Trim(gexDetalleRapport.Value(gexDetalleRapport.Columns("des_color").Index))
    TxtPorcentaje.Text = Trim(gexDetalleRapport.Value(gexDetalleRapport.Columns("porcentaje").Index))
    TxtNroPasadas.Text = Trim(gexDetalleRapport.Value(gexDetalleRapport.Columns("nro_pasadas").Index))
    TxtObservacion.Text = Trim(gexDetalleRapport.Value(gexDetalleRapport.Columns("observaciones").Index))
    
    Call BuscaCombo(Trim(gexDetalleRapport.Value(gexDetalleRapport.Columns("Hilado").Index)), 1, cboCod_HilTel)
    TxtCod_HiladoEstructurado.Text = Trim(gexDetalleRapport.Value(gexDetalleRapport.Columns("cod_hiltel").Index))
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim varcod_hiltel As String
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            LIMPIA_DATOS
            Me.txtSecuencia.Enabled = True
            Me.txtSecuencia.SetFocus
            strSQL = "select cod_hiltel from tx_hilostel where cod_tela='" & Tela & "'"
            varcod_hiltel = DevuelveCampo(strSQL, cCONNECT)
            strSQL = "select cod_hilado_estructurado from it_hilado b where cod_hiltel = '" & varcod_hiltel & "'"
            txtCod_Hiltel.Text = DevuelveCampo(strSQL, cCONNECT)
            TxtDes_Hiltel.Text = DevuelveCampo("SELECT DES_HILTEL FROM IT_HILADO WHERE COD_HILADO_ESTRUCTURADO='" & txtCod_Hiltel.Text & "'", cCONNECT)
            TxtHiloAntiguo = DevuelveCampo("select cod_hiltel from it_hilado where cod_hilado_estructurado='" & txtCod_Hiltel.Text & "'", cCONNECT)
            
            HABILITA
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Case "MODIFICAR"
            If gexDetalleRapport.RowCount = 0 Then Exit Sub
            sTipo = "U"
            HABILITA
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Case "ELIMINAR"
            
            If gexDetalleRapport.RowCount = 0 Then Exit Sub
            vMessage = MsgBox("Esta seguro que desea eliminar el registro", vbYesNo, "Eliminar")
            If vMessage = vbYes Then
                sTipo = "D"
                SALVAR_DATOS
            End If
            CARGA_GRID
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Case "GRABAR"
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                If VALIDA = False Then Exit Sub
                SALVAR_DATOS
                CARGA_GRID
                
                sTipo = ""
                'Call gexDetalleRapport.Find(gexDetalleRapport.Columns("cod_proveedor").Index, jgexEqual, vCod_ProvFind)
        Case "DESHACER"
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            CARGA_GRID
            
        Case "SALIR"
            Unload Me
    End Select
End Sub

Sub SALVAR_DATOS()
Dim vProcedencia As String
On Error GoTo ErrSalvarDatos

'strSQL = "EXEC UP_MAN_TX_RAPPORT_DETALLE '" & sTipo & "'," & _
          txtRapport.Text & ",'" & _
          UCase(txtCod_comb.Text) & "','" & _
          txtSecuencia.Text & "','" & _
          TxtHiloAntiguo.Text & "','" & _
          txtCod_Color.Text & "'," & _
          IIf(TxtNroPasadas.Text = "", 0, TxtNroPasadas.Text) & ",'" & _
          TxtObservacion.Text & "','" & _
          vusu & "','" & Format(Now, "dd/mm/yyyy") & "','" & _
          ComputerName & "'"
    
strSQL = "EXEC UP_MAN_TX_RAPPORT_DETALLE '" & sTipo & "'," & _
          txtRapport.Text & ",'" & _
          UCase(txtCod_comb.Text) & "','" & _
          txtSecuencia.Text & "','" & _
          TxtCod_HiladoEstructurado.Text & "','" & _
          txtCod_Color.Text & "'," & _
          IIf(TxtNroPasadas.Text = "", 0, TxtNroPasadas.Text) & ",'" & _
          TxtObservacion.Text & "','" & _
          vusu & "','" & Format(Now, "dd/mm/yyyy") & "','" & _
          ComputerName & "'"
     
    ExecuteCommandSQL cCONNECT, strSQL
    
Exit Sub
ErrSalvarDatos:
    Screen.MousePointer = vbDefault
    ErrorHandler Err, "SALVAR_DATOS"
End Sub

Sub LIMPIA_DATOS()
    Me.txtSecuencia.Text = ""
    Me.TxtPorcentaje.Text = ""
    txtCod_Hiltel.Text = ""
    TxtDes_Hiltel = ""
    txtCod_Color.Text = ""
    TxtDes_Color.Text = ""
    TxtNroPasadas.Text = ""
    TxtObservacion.Text = ""
    TxtHiloAntiguo.Text = ""
    TxtCod_HiladoEstructurado = ""
    cboCod_HilTel.ListIndex = -1
End Sub


Sub DESHABILITA()
txtSecuencia.Enabled = False
txtCod_Hiltel.Enabled = False
TxtDes_Hiltel.Enabled = False
txtCod_Color.Enabled = False
TxtDes_Color.Enabled = False
TxtNroPasadas.Enabled = False
TxtObservacion.Enabled = False
cboCod_HilTel.Enabled = False
End Sub

Sub HABILITA()
txtCod_Hiltel.Enabled = True
TxtDes_Hiltel.Enabled = True
txtCod_Color.Enabled = True
TxtDes_Color.Enabled = True
TxtNroPasadas.Enabled = True
TxtObservacion.Enabled = True
cboCod_HilTel.Enabled = True
End Sub

Private Sub txtCod_Color_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Trim(txtCod_Color.Text) = "" Then
            Call Me.BUSCA_COLOR(3)
        Else
            Call Me.BUSCA_COLOR(1)
        End If
        SendKeys "{TAB}"
End If
End Sub

Private Sub txtCod_Hiltel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Hiltel.Text) = "" Then
            Command1_Click
        Else
            'StrSql = "SELECT DES_HILTEL FROM IT_HILADO WHERE Cod_HILTEL='" & txtCod_Hiltel.Text & "'"
            strSQL = "SELECT DES_HILTEL FROM IT_HILADO WHERE Cod_HILADO_ESTRUCTURADO='" & txtCod_Hiltel.Text & "'"
            Me.TxtDes_Hiltel.Text = DevuelveCampo(strSQL, cCONNECT)
            strSQL = "select cod_hiltel from it_hilado where cod_hilado_estructurado='" & txtCod_Hiltel.Text & "'"
            TxtHiloAntiguo.Text = DevuelveCampo(strSQL, cCONNECT)
        End If
    SendKeys "{TAB}"
    End If
End Sub

Private Sub TxtDes_Color_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Me.BUSCA_COLOR(2)
    'SendKeys "{TAB}"
    TxtNroPasadas.SetFocus
End If

End Sub

Private Sub TxtDes_Hiltel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        busca_hilo_des
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtNroPasadas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtObservacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtSecuencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim oTipo As New frmBusqGeneral
    Dim Rs As New ADODB.Recordset
    
    Set oTipo.oParent = Me
    strSQL = "SELECT secuencia as Secuencia, Porcentaje as Porcentaje " & _
            " From tx_rapport_composicion " & _
            " Where rapport_number = " & Trim(txtRapport) & _
            " AND SECUENCIA NOT IN (SELECT secuencia From tx_rapport_DETALLE " & _
            " where rapport_number= " & Trim(txtRapport) & " AND rapport_COMB = '" & Trim(Me.txtCod_comb) & "')"

    oTipo.sQuery = strSQL
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtSecuencia.Text = Trim(Codigo)
        TxtPorcentaje.Text = Trim(Descripcion)
    Else
        txtSecuencia.Text = ""
        TxtPorcentaje.Text = ""
    End If
    Codigo = ""
    Descripcion = ""
    Set oTipo = Nothing
    Set Rs = Nothing
    SendKeys "{TAB}"
    
End If
End Sub

Private Sub txtSecuencia_LostFocus()
If txtSecuencia <> "" Then
    TxtPorcentaje = DevuelveCampo("SELECT PORCENTAJE FROM TX_RAPPORT_COMPOSICION WHERE RAPPORT_NUMBER=" & txtRapport & " AND SECUENCIA='" & txtSecuencia & "'", cCONNECT)
End If
End Sub

Public Sub BUSCA_COLOR(tipo As Integer)
    Select Case tipo
        Case 1:
                    strSQL = "SELECT DES_COLOR  FROM LB_COLOR WHERE COD_COLOR = '" & txtCod_Color & "'"
                    Me.TxtDes_Color.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                    If Trim(TxtDes_Color.Text) <> "" Then SendKeys "{TAB}", True
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim Rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        oTipo.sQuery = "SELECT Cod_Color as 'Código', Des_Color as 'Descripción' FROM LB_COLOR WHERE Des_COLOR LIKE '%" & Trim(TxtDes_Color.Text) & "%' ORDER BY des_COLOR"
                    Else
                        oTipo.sQuery = "SELECT Cod_color as 'Código', Des_color as 'Descripción' FROM LB_COLOR ORDER BY des_COLOR"
                    End If
                    
                    oTipo.Cargar_Datos
                    oTipo.Show 1
                    If Codigo <> "" Then
                        Me.txtCod_Color = Trim(Codigo)
                        Me.TxtDes_Color.Text = Trim(Descripcion)
                        Codigo = "": Descripcion = ""
                    Else
                        Me.txtCod_Color = ""
                        Me.TxtDes_Color.Text = ""
                    End If
                    Set oTipo = Nothing
                    Set Rs = Nothing
    End Select
    
End Sub
Function VALIDA() As Boolean
    If Trim(txtSecuencia) = "" Then
        MsgBox "Ingrese Secuencia"
        VALIDA = False
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Exit Function
    End If
    If Trim(TxtNroPasadas.Text) = "" Then
        MsgBox "Ingrese Nro. de pasadas"
        VALIDA = False
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Exit Function
    End If
    If Trim(txtCod_Hiltel.Text) = "" Then
        MsgBox "Ingrese Hilado"
        VALIDA = False
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Exit Function
    End If
    VALIDA = True
End Function


