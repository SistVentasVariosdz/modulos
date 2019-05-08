VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmMantTelaCombDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Combinaciones"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   12075
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Hilados"
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   825
      TabIndex        =   10
      Top             =   6285
      Width           =   1965
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmMantTelaCombDet.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ultimo"
         Top             =   75
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "frmMantTelaCombDet.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Siguiente"
         Top             =   75
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmMantTelaCombDet.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Anterior"
         Top             =   75
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   15
         Picture         =   "frmMantTelaCombDet.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Primero"
         Top             =   75
         Width           =   495
      End
   End
   Begin VB.Frame Fralista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Left            =   90
      TabIndex        =   7
      Tag             =   "List"
      Top             =   105
      Width           =   11805
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2925
         Left            =   240
         TabIndex        =   8
         Top             =   255
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   5159
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   17
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
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
      Height          =   2835
      Left            =   120
      TabIndex        =   0
      Tag             =   "Detail"
      Top             =   3405
      Width           =   11835
      Begin VB.Frame FraDesarrollo 
         Height          =   1095
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   7575
         Begin VB.TextBox txtLong_Malla 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5280
            TabIndex        =   24
            Text            =   "0"
            Top             =   270
            Width           =   900
         End
         Begin VB.TextBox txtNom_Malla 
            Height          =   285
            Left            =   1170
            MaxLength       =   50
            TabIndex        =   23
            Top             =   240
            Width           =   1830
         End
         Begin VB.TextBox TxtNum_Agujas 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   25
            Text            =   "0"
            Top             =   645
            Width           =   780
         End
         Begin VB.TextBox TxtNum_Alimentadores 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5280
            TabIndex        =   26
            Text            =   "0"
            Top             =   645
            Width           =   780
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Malla"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Longitud Malla :"
            Height          =   180
            Left            =   3720
            TabIndex        =   31
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Num. Agujas :"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   735
            Width           =   990
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Num. Alimentadores :"
            Height          =   195
            Left            =   3720
            TabIndex        =   29
            Top             =   735
            Width           =   1500
         End
      End
      Begin VB.TextBox TxtCod_HiladoEstructurado 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1350
         TabIndex        =   22
         Top             =   1320
         Width           =   1515
      End
      Begin VB.ComboBox cboCod_HilTel 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   10365
      End
      Begin VB.TextBox TxtHiloNuevo 
         BackColor       =   &H80000009&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   20
         Top             =   1350
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox txtDes_Color 
         Height          =   285
         Left            =   2940
         MaxLength       =   50
         TabIndex        =   6
         Top             =   960
         Width           =   3720
      End
      Begin VB.TextBox txtDes_HilTel 
         Height          =   285
         Left            =   6540
         TabIndex        =   4
         Top             =   1350
         Visible         =   0   'False
         Width           =   3720
      End
      Begin VB.CommandButton cmdBusColor 
         Caption         =   "..."
         Height          =   330
         Left            =   2540
         TabIndex        =   19
         Tag             =   "..."
         Top             =   945
         Width           =   360
      End
      Begin VB.CommandButton cmdBusHilo 
         Caption         =   "..."
         Height          =   330
         Left            =   6135
         TabIndex        =   18
         Tag             =   "..."
         Top             =   1320
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox txtCod_HilTel 
         BackColor       =   &H80000009&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7200
         TabIndex        =   3
         Top             =   255
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cboNum_Secuencia 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtCod_Color 
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   5
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Hilado Antiguo"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1365
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hilo Nuevo"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   705
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Hilo Antiguo"
         Height          =   195
         Left            =   6120
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Color"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1050
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Secuencia :"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   345
         Width           =   855
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   3360
      TabIndex        =   9
      Top             =   6315
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantTelaCombDet.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantTelaCombDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public Codigo, Descripcion As String
Dim sTipo As String
Dim Rs_Grid As New ADODB.Recordset
Dim StrSql As String
Public Codigo_tela As String, sCod_Cliente As String
Public Codigo_Comb As String, sCod_TemCli As String
Public Rapport_Number As String
Public Rapport_Comb As String
Public Cod_FamTela As String

Private Sub cboCod_HilTel_Click()
TxtCod_HiladoEstructurado = Trim(Right(cboCod_HilTel.Text, 10))
End Sub

Private Sub cboNum_Secuencia_Click()
If sTipo = "I" And cboNum_Secuencia.ListIndex <> -1 Then
    StrSql = "SELECT Cod_HilTel FROM tx_hilostel WHERE Cod_Tela ='" & Codigo_tela & "' AND Num_Secuencia='" & Mid(cboNum_Secuencia.Text, 1, 2) & "'"
    txtCod_HilTel.Text = Trim(DevuelveCampo(StrSql, cCONNECT))
'    If Trim(txtCod_HilTel.Text) <> "" Then
'        Call txtCod_HilTel_KeyPress(13)
'    End If
    StrSql = "SELECT cod_hilado_estructurado FROM it_hilado WHERE Cod_hiltel ='" & txtCod_HilTel.Text & "'"
    TxtHiloNuevo.Text = DevuelveCampo(StrSql, cCONNECT)
    If Trim(TxtHiloNuevo.Text) <> "" Then
        Call txtCod_HilTel_KeyPress(13)
    End If
End If
End Sub

Private Sub cmdBusColor_Click()
Dim oTipo As New frmBusqGeneral
Dim Rs As New ADODB.Recordset
    Load oTipo
    Set oTipo.oParent = Me
    'oTipo.sQuery = "SELECT Cod_Color as Código, Des_Color as Descripción FROM LB_Color"
    oTipo.sQuery = "EXEC SM_MUESTRA_COLORES_HILADO_SEGUN_CARTA '" & sCod_Cliente _
                   & "', '" & sCod_TemCli & "'"
    oTipo.iCampo = 1
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtCod_Color.Text = Codigo
        txtDes_Color.Text = Descripcion
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub

Private Sub cmdBusHilo_Click()
    Dim oTipo As New frmBusqGeneral
    Dim Rs As New ADODB.Recordset
    Load oTipo
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT cod_hilado_estructurado as Codigo, des_HilTel as Descripcion FROM IT_Hilado WHERE Cod_Hilado_estructurado<>''"
    oTipo.iCampo = 1
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        TxtHiloNuevo.Text = Codigo
        txtDes_HilTel.Text = Descripcion
        txtCod_HilTel.Text = DevuelveCampo("SELECT COD_HILTEL FROM IT_HILADO WHERE COD_HILADO_ESTRUCTURADO='" & Codigo & "'", cCONNECT)
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub

Private Sub cmdFirst_Click()
    If Not Rs_Grid.BOF Then
        Rs_Grid.MoveFirst
    End If
End Sub
Private Sub cmdLast_Click()
    If Not Rs_Grid.EOF Then
        Rs_Grid.MoveLast
    End If
End Sub
Private Sub cmdNext_Click()
    If Not Rs_Grid.EOF Then
        Rs_Grid.MoveNext
    End If
End Sub
Private Sub cmdPrevious_Click()
    If Not Rs_Grid.BOF Then
        Rs_Grid.MovePrevious
    End If
End Sub

Private Sub Form_Load()
    Call FormSet(Me)
    FormateaGrid Me.DGridLista
    HabilitaMant Me.MantFunc1, ""
    DESHABILITA_DATOS
    Call CargaCombos
    MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub
Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Rs_Grid.State <> 1 Then
        Exit Sub
    End If
    If Not Rs_Grid.EOF And Not Rs_Grid.BOF Then
        Call Carga_Datos
        DESHABILITA_DATOS
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Rs_Grid = Nothing
End Sub

Private Sub FunctDetalles_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
    
    Select Case ActionName
        Case "ADICIONAR"
            If FixNulos(RTrim(Rapport_Number), vbLong) = 0 Then
                sTipo = "I"
                LIMPIAR_DATOS
                HABILITA_DATOS
                HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
                DGridLista.Enabled = False
                cboNum_Secuencia.SetFocus
            End If
        Case "MODIFICAR"
            If FixNulos(RTrim(Rapport_Number), vbLong) = 0 Then
                sTipo = "U"
                HABILITA_DATOS
                cboNum_Secuencia.Enabled = False
                HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
                DGridLista.Enabled = False
                txtCod_Color.SetFocus
            End If
        Case "ELIMINAR"
            If FixNulos(RTrim(Rapport_Number), vbLong) = 0 Then
                Eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Combinación-Detalle")
                If Eliminar = vbYes Then
                    sTipo = "D"
                    ELIMINAR_DATOS
                    RECARGAR_DATOS
                End If
            End If
        Case "GRABAR"
            If FixNulos(RTrim(Rapport_Number), vbLong) = 0 Then
                If VALIDA_DATOS Then
                    SALVAR_DATOS
                    RECARGAR_DATOS
                    HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                    DGridLista.Enabled = True
                    sTipo = ""
                End If
            End If
        Case "DESHACER"
            If FixNulos(RTrim(Rapport_Number), vbLong) = 0 Then
                LIMPIAR_DATOS
                RECARGAR_DATOS
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                DGridLista.Enabled = True
                sTipo = ""
            End If
         Case "SALIR"
            Unload Me
    End Select
End Sub

Sub LIMPIAR_DATOS()
    cboNum_Secuencia.ListIndex = -1
    txtCod_Color.Text = ""
    txtDes_Color.Text = ""
    txtCod_HilTel.Text = ""
    txtDes_HilTel.Text = ""
    TxtHiloNuevo.Text = ""
    
    TxtCod_HiladoEstructurado = ""
    cboCod_HilTel.ListIndex = -1
    txtLong_Malla.Text = 0
    txtNom_Malla.Text = ""
    TxtNum_Agujas.Text = 0
    TxtNum_Alimentadores.Text = 0
End Sub


Public Sub CargaCombos()
   
    'Llena Combo de Secuencias
    StrSql = "SELECT Num_Secuencia + '     ' + CONVERT(VARCHAR,Porcentaje ) +  '[%]'  FROM tx_hilostel WHERE Cod_Tela ='" & Codigo_tela & "'"
    Call LlenaCombo(cboNum_Secuencia, StrSql, cCONNECT)
    
    StrSql = "SM_Muestra_It_Hilado"
    Call LlenaCombo(cboCod_HilTel, StrSql, cCONNECT)
    
End Sub

Sub Carga_Datos()
Dim intCont As Integer
Dim strTexto As String

    If Not Rs_Grid.EOF Then
       
        'txtNum_Secuencia.Text = Trim(Rs_Grid("Num_Secuencia").Value)
        
        txtCod_Color.Text = Trim(Rs_Grid("Cod_Color").Value)
        'Call txtCod_Color_KeyPress(13)
        
        StrSql = "SELECT Des_Color FROM LB_Color WHERE Cod_Color='" & txtCod_Color.Text & "'"
        txtDes_Color.Text = DevuelveCampo(StrSql, cCONNECT)
        
        
        Call BuscaCombo(Rs_Grid("Num_Secuencia").Value & "     " & CStr(Format(Rs_Grid("Porcentaje").Value, "0.00")) & "[%]", 2, cboNum_Secuencia)
        
        'Ant. 30/06
        'txtCod_HilTel.Text = Trim(Rs_Grid("Hilado").Value)
        'Call txtCod_HilTel_KeyPress(13)
        'StrSql = "SELECT cod_hilado_estructurado FROM it_hilado WHERE Cod_hiltel='" & Trim(Rs_Grid("Hilado").Value) & "'"
        'TxtHiloNuevo.Text = DevuelveCampo(StrSql, cCONNECT)
                
        'Pos. 30/06
        Call BuscaCombo(IIf(IsNull(Rs_Grid("cod_hilado_estructurado")), "", Rs_Grid("cod_hilado_estructurado")), 1, cboCod_HilTel)
        TxtCod_HiladoEstructurado.Text = Trim(Rs_Grid("Hilado"))
        If UCase(Cod_FamTela) = "DE" Then
            txtLong_Malla.Text = Trim(Rs_Grid("long_malla"))
            txtNom_Malla.Text = Trim(Rs_Grid("nom_malla"))
            TxtNum_Agujas.Text = Trim(Rs_Grid("num_agujas"))
            TxtNum_Alimentadores.Text = Trim(Rs_Grid("num_alimentadores"))
        End If
        
    End If

End Sub

Sub RECARGAR_DATOS()
    
    Rs_Grid.Close
    CARGA_GRID
    
End Sub

Public Sub CARGA_GRID()
    Dim StrSql As String
    Set Rs_Grid = New ADODB.Recordset
    Rs_Grid.ActiveConnection = cCONNECT
    Rs_Grid.CursorType = adOpenStatic
    Rs_Grid.CursorLocation = adUseClient
    Rs_Grid.LockType = adLockReadOnly
    
    'Esta cadena es para devolver el Codigo de Cliente
    'StrSql = "EXEC UP_SEL_TELACOMBDET '" & Codigo_tela & "','" & Codigo_Comb & "'"
    
    If FixNulos(Rapport_Number, vbLong) = 0 Then
        StrSql = "EXEC SM_MUESTRA_DETALLE_TX_TELACOMBDET '" & Codigo_tela & "','" & Codigo_Comb & "'"
    Else
        StrSql = "EXEC SM_MUESTRA_TX_RAPPORT_DETALLE '" & Rapport_Number & "','" & Rapport_Comb & "'"
    End If
    
    Rs_Grid.Open StrSql
    Set DGridLista.DataSource = Rs_Grid
    DGridLista.Refresh

    If Rs_Grid.RecordCount > 0 Then
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Call Carga_Datos
    Else
        HabilitaMant Me.MantFunc1, "ADICIONAR"
        Call LIMPIAR_DATOS
    End If
End Sub

Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    
    If sTipo = "I" Then
        If cboNum_Secuencia.Text = "" Then
            MsgBox ("Usted no ha seleccionado un Numero de Secuencia. Sirvase Verificar")
            VALIDA_DATOS = False
            Exit Function
        End If
    
        StrSql = "SELECT Num_Secuencia FROM TX_TELACOMBDET WHERE Cod_Tela='" & Codigo_tela & "' AND Cod_Comb='" & Codigo_Comb & "' AND Num_Secuencia='" & Mid(cboNum_Secuencia.Text, 1, 2) & "'"
        If DevuelveCampo(StrSql, cCONNECT) <> "" Then
            MsgBox ("Ese registro ya se encuentra ingresado. Sirvase verificar")
            VALIDA_DATOS = False
            Exit Function
        End If
    End If
    
    If sTipo <> "D" Then
        'Verifica Color del Hilado
'        If Trim(txtCod_Color.Text) <> "" Then
'            strSQL = "SELECT Cod_Color FROM LB_Color WHERE Cod_Color='" & txtCod_Color.Text & "'"
'            If DevuelveCampo(strSQL, cCONNECT) = "" Then
'                MsgBox ("El código de color ingresado no existe, Sirvase verficar")
'                    VALIDA_DATOS = False
'                txtCod_Color.SetFocus
'            End If
'        End If
        
        If Trim(txtCod_HilTel.Text) <> "" Then
            StrSql = "SELECT Cod_HilTel, des_HilTel FROM IT_Hilado WHERE Cod_HilTel='" & Trim(txtCod_HilTel.Text) & "'"
            If DevuelveCampo(StrSql, cCONNECT) = "" Then
                MsgBox ("El código de hilo ingresado no existe, Sirvase verficar")
                    VALIDA_DATOS = False
                    txtCod_HilTel.SetFocus
            End If
        End If
    End If
End Function

Sub HABILITA_DATOS()
   
    cboNum_Secuencia.Enabled = True
    txtCod_Color.Enabled = True
    txtDes_Color.Enabled = True
    'txtCod_HilTel.Enabled = True
    txtDes_HilTel.Enabled = True
    cmdBusHilo.Enabled = True
    cmdBusColor.Enabled = True
    TxtHiloNuevo.Enabled = True
    cboCod_HilTel.Enabled = True
    FraDesarrollo.Enabled = True
End Sub

Sub DESHABILITA_DATOS()

    'txtDes_Comb.Enabled = False
    cboNum_Secuencia.Enabled = False
    txtCod_Color.Enabled = False
    txtDes_Color.Enabled = False
    'txtCod_HilTel.Enabled = False
    txtDes_HilTel.Enabled = False
    cmdBusHilo.Enabled = False
    cmdBusColor.Enabled = False
    TxtHiloNuevo.Enabled = False
    cboCod_HilTel.Enabled = False
    FraDesarrollo.Enabled = False
End Sub

Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim StrSql As String
    
    Con.ConnectionString = cCONNECT
    Con.Open
    
        Con.BeginTrans

        'StrSql = "EXEC UP_MAN_TELACOMBDET '" & _
        sTipo & "','" & _
        Codigo_tela & "','" & _
        Codigo_Comb & "','" & _
        Mid(cboNum_Secuencia.Text, 1, 2) & "','" & _
        txtCod_Color.Text & "','" & _
        txtCod_HilTel.Text & "','" & vusu & "'"
        
        StrSql = "EXEC UP_MAN_TELACOMBDET '" & _
        sTipo & "','" & _
        Codigo_tela & "','" & _
        Codigo_Comb & "','" & _
        Mid(cboNum_Secuencia.Text, 1, 2) & "','" & _
        txtCod_Color.Text & "','" & _
        TxtCod_HiladoEstructurado & "','" & vusu & "','" & txtLong_Malla.Text & "','" & txtNom_Malla.Text & "'," & Val(TxtNum_Agujas) & "," & Val(TxtNum_Alimentadores)
       
        
        Con.Execute StrSql

        Con.CommitTrans
        Dim amensaje As New clsMessages
        amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
        Informa "", amensaje
        
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub

Sub ELIMINAR_DATOS()
 Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
   
    Con.ConnectionString = cCONNECT
    Con.Open
    Con.BeginTrans
       
        'StrSql = "EXEC UP_MAN_TELACOMBDET '" & _
        sTipo & "','" & _
        Codigo_tela & "','" & _
        Codigo_Comb & "','" & _
        Mid(cboNum_Secuencia.Text, 1, 2) & "','" & _
        txtCod_Color.Text & "','" & _
        txtCod_HilTel.Text & "','" & vusu & "'"
        
        StrSql = "EXEC UP_MAN_TELACOMBDET '" & _
        sTipo & "','" & _
        Codigo_tela & "','" & _
        Codigo_Comb & "','" & _
        Mid(cboNum_Secuencia.Text, 1, 2) & "','" & _
        txtCod_Color.Text & "','" & _
        TxtCod_HiladoEstructurado & "','" & vusu & "'"
        
        Con.Execute StrSql
        
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

Private Sub txtCod_Color_KeyPress(KeyAscii As Integer)
    Dim StrSql As String
    If KeyAscii = 13 Then
        If Trim(txtCod_Color.Text) = "" Then
            cmdBusColor_Click
        Else
            StrSql = "SELECT Des_Color FROM LB_Color WHERE Cod_Color='" & txtCod_Color.Text & "'"
            txtDes_Color.Text = DevuelveCampo(StrSql, cCONNECT)
        End If
    End If
End Sub

Private Sub txtCod_HilTel_KeyPress(KeyAscii As Integer)
    Dim StrSql As String
    If KeyAscii = 13 Then
        If Trim(txtCod_HilTel.Text) = "" Then
            cmdBusHilo_Click
        Else
            StrSql = "SELECT des_HilTel FROM IT_Hilado WHERE Cod_HilTel='" & txtCod_HilTel.Text & "'"
            txtDes_HilTel.Text = DevuelveCampo(StrSql, cCONNECT)
        End If
    End If
End Sub

Private Sub txtDes_Color_KeyPress(KeyAscii As Integer)
    Dim StrSql As String
    If KeyAscii = 13 Then
    
        Dim oTipo As New frmBusqGeneral
        Dim Rs As New ADODB.Recordset
        Set oTipo.oParent = Me
        oTipo.sQuery = "SELECT Cod_Color as Código, Des_Color as Descripción FROM LB_Color WHERE Des_Color LIKE '" & txtDes_Color.Text & "%'"
        oTipo.Cargar_Datos
        oTipo.Show 1
        If Codigo <> "" Then
            txtCod_Color.Text = Codigo
            txtDes_Color.Text = Descripcion
        End If
        Set oTipo = Nothing
        Set Rs = Nothing
    
    
        'If Len(txtDes_Color.Text) < 5 Then
        '    Call MsgBox("La descripción debe tener como mínimo 5 caracteres. Sirvase verificar", vbInformation)
        '    Exit Sub
        'Else
            'strSQL = "SELECT Cod_Color FROM LB_Color WHERE Des_Color LIKE '" & txtCod_Color.Text & "%'"
            'txtCod_Color.Text = DevuelveCampo(strSQL, cCONNECT)
        'End If
    End If
End Sub


Private Sub txtDes_HilTel_KeyPress(KeyAscii As Integer)
    Dim StrSql As String
    If KeyAscii = 13 Then
        If Len(txtDes_HilTel.Text) < 5 Then
            Call MsgBox("La descripción debe tener como mínimo 5 caracteres. Sirvase verificar", vbInformation)
            Exit Sub
        Else
            StrSql = "SELECT Cod_HilTel FROM IT_Hilado WHERE Des_HilTel LIKE '" & txtDes_HilTel.Text & "%'"
            txtCod_HilTel.Text = DevuelveCampo(StrSql, cCONNECT)
            StrSql = "SELECT cod_hilado_estructurado FROM IT_Hilado WHERE des_HilTel LIKE '" & txtDes_HilTel.Text & "%'"
            TxtHiloNuevo.Text = DevuelveCampo(StrSql, cCONNECT)
        End If
    End If
End Sub

Private Sub TxtHiloNuevo_KeyPress(KeyAscii As Integer)
    Dim StrSql As String
    If KeyAscii = 13 Then
        If Trim(TxtHiloNuevo.Text) = "" Then
            cmdBusHilo_Click
            StrSql = "SELECT cod_HilTel FROM IT_Hilado WHERE cod_hilado_estructurado='" & TxtHiloNuevo.Text & "'"
            txtCod_HilTel.Text = Trim(DevuelveCampo(StrSql, cCONNECT))
        Else
            StrSql = "SELECT des_HilTel FROM IT_Hilado WHERE cod_hilado_estructurado='" & TxtHiloNuevo.Text & "'"
            txtDes_HilTel.Text = Trim(DevuelveCampo(StrSql, cCONNECT))
            StrSql = "SELECT cod_HilTel FROM IT_Hilado WHERE cod_hilado_estructurado='" & TxtHiloNuevo.Text & "'"
            txtCod_HilTel.Text = Trim(DevuelveCampo(StrSql, cCONNECT))
        End If
    End If
End Sub

Private Sub txtLong_Malla_GotFocus()
SelectionText txtLong_Malla
End Sub

Private Sub txtLong_Malla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(txtLong_Malla.Text) = "" Then txtLong_Malla.Text = 0
    SendKeys "{TAB}"
Else
    Call SoloNumeros(txtLong_Malla, KeyAscii, True, 3)
End If
End Sub

Private Sub txtNom_Malla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtNum_Agujas_GotFocus()
SelectionText TxtNum_Agujas
End Sub

Private Sub TxtNum_Agujas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtNum_Agujas, KeyAscii, False)
End If
End Sub

Private Sub TxtNum_Alimentadores_GotFocus()
SelectionText TxtNum_Alimentadores
End Sub

Private Sub TxtNum_Alimentadores_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtNum_Alimentadores, KeyAscii, False)
End If
End Sub
