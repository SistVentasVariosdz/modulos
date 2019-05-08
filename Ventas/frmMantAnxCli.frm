VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Begin VB.Form frmMantAnxCli 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anexos Contables x Cliente"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   1890
      TabIndex        =   9
      Top             =   4485
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantAnxCli.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin VB.TextBox txtCod_TipAnex 
      Height          =   315
      Left            =   1380
      MaxLength       =   1
      TabIndex        =   4
      Top             =   3900
      Width           =   375
   End
   Begin VB.TextBox txtDes_Anexo 
      Height          =   315
      Left            =   2550
      TabIndex        =   6
      Top             =   3900
      Width           =   4485
   End
   Begin VB.TextBox txtCod_Anxo 
      Height          =   315
      Left            =   1755
      TabIndex        =   5
      Top             =   3900
      Width           =   795
   End
   Begin GridEX20.GridEX gexAnx 
      Height          =   3045
      Left            =   60
      TabIndex        =   3
      Top             =   585
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5371
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmMantAnxCli.frx":0160
      Column(2)       =   "frmMantAnxCli.frx":0228
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmMantAnxCli.frx":02CC
      FormatStyle(2)  =   "frmMantAnxCli.frx":0404
      FormatStyle(3)  =   "frmMantAnxCli.frx":04B4
      FormatStyle(4)  =   "frmMantAnxCli.frx":0568
      FormatStyle(5)  =   "frmMantAnxCli.frx":0640
      FormatStyle(6)  =   "frmMantAnxCli.frx":06F8
      ImageCount      =   0
      PrinterProperties=   "frmMantAnxCli.frx":07D8
   End
   Begin FunctionsButtons.FunctButt fnbBuscar 
      Height          =   510
      Left            =   6330
      TabIndex        =   2
      Top             =   45
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~BUSCAR~True~True~&Buscar~0~0~1~~0~False~False~&Buscar~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.TextBox txtNom_Cliente 
      Height          =   315
      Left            =   1695
      TabIndex        =   1
      Top             =   180
      Width           =   4485
   End
   Begin VB.TextBox txtAbr_Cliente 
      Height          =   315
      Left            =   885
      TabIndex        =   0
      Top             =   180
      Width           =   795
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   6015
      Top             =   4650
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "Anexo Contable"
      Height          =   405
      Left            =   465
      TabIndex        =   8
      Top             =   3855
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   285
      Left            =   105
      TabIndex        =   7
      Top             =   195
      Width           =   735
   End
End
Attribute VB_Name = "frmMantAnxCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo As String, Descripcion As String, TIpoAdd As String
Dim sTit As String, sErr As String, StrSql As String, sAccion As String, rstAux As ADODB.Recordset

Private Sub fnbBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo ErrBusq
    
    sTit = "Muestra Anexos Permitidos por Cliente"
    
    Screen.MousePointer = 11
    
    StrSql = "EXEC TG_UP_MAN_CLIENTE_ANEXOCONT 'S', '" & txtAbr_Cliente.Tag & "', '" & _
             txtCod_TipAnex & "', '" & txtCod_Anxo & "'"
    Set gexAnx.ADORecordset = CargarRecordSetDesconectado(StrSql, cCONNECT)
    
    gexAnx.Columns("Cod_Cliente").Width = 105
    gexAnx.Columns("Nom_Cliente").Width = 90
    gexAnx.Columns("Cod_TipAnex").Width = 570
    gexAnx.Columns("Cod_Anxo").Width = 1110
    gexAnx.Columns("Des_Anexo").Width = 3480
    
    gexAnx.Columns("Cod_Cliente").Visible = False
    gexAnx.Columns("Nom_Cliente").Visible = False
    gexAnx.Columns("Cod_TipAnex").Caption = "Tipo"
    gexAnx.Columns("Cod_Anxo").Caption = "Codigo"
    gexAnx.Columns("Des_Anexo").Caption = "Anexo"
    
    MantFunc1.FunctionsUser = "ADICIONAR/ELIMINAR"
    Habilita False
    gexAnx.Enabled = True
    
    Screen.MousePointer = 0
Exit Sub
ErrBusq:
    sErr = Err.Description
    Screen.MousePointer = 0
    MsgBox sErr, vbCritical + vbOKOnly, sTit
End Sub

Private Sub Form_Load()
    fnbBuscar_ActionClick 0, 0, ""
End Sub

Public Sub BuscaCliente(Opcion As String)
    
    StrSql = "SELECT Cod_Cliente, Abr_Cliente, Nom_Cliente FROM TG_CLIENTE WHERE "
    
    txtAbr_Cliente = Trim(txtAbr_Cliente)
    txtNom_Cliente = Trim(txtNom_Cliente)
    
    Select Case Opcion
    Case 1: StrSql = StrSql & "Abr_Cliente LIKE '%" & txtAbr_Cliente & "%'"
    Case 2: StrSql = StrSql & "Nom_Cliente LIKE '%" & txtNom_Cliente & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = StrSql
    frmBusqGeneral3.Cargar_Datos
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    'frmBusqGeneralJanus.Show vbModal
    
    frmBusqGeneral3.gexLista.Columns("Cod_Cliente").Visible = False
    frmBusqGeneral3.gexLista.Columns("Abr_Cliente").Width = 570
    frmBusqGeneral3.gexLista.Columns("Nom_Cliente").Width = 2370
    
    frmBusqGeneral3.gexLista.Columns("Abr_Cliente").Caption = "Abrev."
    frmBusqGeneral3.gexLista.Columns("Nom_Cliente").Caption = "Cliente"
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtAbr_Cliente.Tag = ""
    txtAbr_Cliente = ""
    txtNom_Cliente = ""
    
    If Codigo <> "" Then
        txtAbr_Cliente.Tag = rstAux!Cod_Cliente
        txtAbr_Cliente = rstAux!Abr_Cliente
        txtNom_Cliente = rstAux!Nom_Cliente
        
    End If
    Codigo = ""
    Descripcion = ""
End Sub

Private Sub gexAnx_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Limpiar
    If gexAnx.Row > 0 Then
        txtCod_TipAnex = gexAnx.Value(gexAnx.Columns("Cod_TipAnex").Index)
        txtDes_Anexo = gexAnx.Value(gexAnx.Columns("Des_Anexo").Index)
        txtCod_Anxo = gexAnx.Value(gexAnx.Columns("Cod_Anxo").Index)
        txtAbr_Cliente.Tag = gexAnx.Value(gexAnx.Columns("Cod_Cliente").Index)
        txtAbr_Cliente = gexAnx.Value(gexAnx.Columns("Abr_Cliente").Index)
        txtNom_Cliente = gexAnx.Value(gexAnx.Columns("Nom_Cliente").Index)
    End If
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "ADICIONAR"
        sAccion = "I"
        gexAnx.Enabled = False
        Limpiar
        Habilita True
        txtCod_TipAnex.SetFocus
        MantFunc1.FunctionsUser = "GRABAR/DESHACER"
    Case "MODIFICAR"
        'Nada
    Case "ELIMINAR"
        sAccion = "D"
        If Not GuardarDatos Then Exit Sub
        fnbBuscar_ActionClick 0, 0, ""
    Case "GRABAR"
        If Not GuardarDatos Then Exit Sub
        gexAnx.Enabled = True
        Habilita False
        MantFunc1.FunctionsUser = "ADICIONAR/ELIMINAR"
        fnbBuscar_ActionClick 0, 0, ""
    Case "DESHACER"
        sAccion = ""
        gexAnx.Enabled = True
        Habilita False
        MantFunc1.FunctionsUser = "ADICIONAR/ELIMINAR"
    Case "SALIR"
        Unload Me
    End Select
End Sub

Private Function GuardarDatos() As Boolean
On Error GoTo ErrDatos
    
    GuardarDatos = False
    
    sTit = "Guardar Datos"
    
    Screen.MousePointer = 11
    
    StrSql = "EXEC TG_UP_MAN_CLIENTE_ANEXOCONT '" & sAccion & "', '" & _
             txtAbr_Cliente.Tag & "', '" & txtCod_TipAnex & "', '" & txtCod_Anxo & "'"
    ExecuteCommandSQL cCONNECT, StrSql
    
    Screen.MousePointer = 0
    
    GuardarDatos = True
    
Exit Function
ErrDatos:
    sErr = Err.Description
    MsgBox sErr, vbCritical + vbOKOnly, sTit
End Function

Private Sub Limpiar()
    txtCod_TipAnex = ""
    txtDes_Anexo = ""
    txtCod_Anxo = ""
End Sub

Private Sub Habilita(Modo As Boolean)
    txtCod_TipAnex.Enabled = Modo
    txtDes_Anexo.Enabled = Modo
    txtCod_Anxo.Enabled = Modo
End Sub

Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaCliente 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_Anxo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaAnexo 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_TipAnex_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Trim(txtCod_Anxo) <> "" Then BuscaAnexo 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDes_Anexo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaAnexo 2
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaCliente 2
        SendKeys "{TAB}"
    End If
End Sub

Private Sub BuscaAnexo(Opcion As String)
    
    StrSql = "SELECT Cod_TipAnex, Cod_Anxo, Des_Anexo FROM CN_ANEXOSCONTABLES " & _
             "WHERE Cod_TipAnex = '" & txtCod_TipAnex & "' AND "
    
    txtCod_Anxo = Trim(txtCod_Anxo)
    txtDes_Anexo = Trim(txtDes_Anexo)
    
    Select Case Opcion
    Case 1: StrSql = StrSql & "Cod_Anxo LIKE '%" & txtCod_Anxo & "%'"
    Case 2: StrSql = StrSql & "Des_Anexo LIKE '%" & txtDes_Anexo & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = StrSql
    frmBusqGeneral3.Cargar_Datos
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    'frmBusqGeneralJanus.Show vbModal
    
    frmBusqGeneral3.gexLista.Columns("Cod_TipAnex").Width = 400
    frmBusqGeneral3.gexLista.Columns("Cod_Anxo").Width = 570
    frmBusqGeneral3.gexLista.Columns("Des_Anexo").Width = 2370
    
    frmBusqGeneral3.gexLista.Columns("Cod_TipAnex").Caption = "Tipo"
    frmBusqGeneral3.gexLista.Columns("Cod_Anxo").Caption = "Codigo"
    frmBusqGeneral3.gexLista.Columns("Des_Anexo").Caption = "Anexo Contable"
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtCod_Anxo = ""
    txtDes_Anexo = ""
    
    If Codigo <> "" Then
        txtCod_TipAnex = rstAux!Cod_TipAnex
        txtCod_Anxo = rstAux!Cod_Anxo
        txtDes_Anexo = Trim(rstAux!Des_Anexo)
    End If
    Codigo = ""
    Descripcion = ""
End Sub

