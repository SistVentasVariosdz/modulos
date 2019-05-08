VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmReqVsRealAvios 
   Caption         =   "Requerimiento Real Avios/Telas"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   440
      Left            =   240
      TabIndex        =   14
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   9240
      TabIndex        =   11
      Top             =   5520
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalle"
      Height          =   4290
      Left            =   60
      TabIndex        =   7
      Top             =   990
      Width           =   10380
      Begin GridEX20.GridEX gexList 
         Height          =   4005
         Left            =   90
         TabIndex        =   5
         Top             =   195
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   7064
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         SelectionStyle  =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmReqVsRealAvios.frx":0000
         Column(2)       =   "frmReqVsRealAvios.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmReqVsRealAvios.frx":016C
         FormatStyle(2)  =   "frmReqVsRealAvios.frx":02A4
         FormatStyle(3)  =   "frmReqVsRealAvios.frx":0354
         FormatStyle(4)  =   "frmReqVsRealAvios.frx":0408
         FormatStyle(5)  =   "frmReqVsRealAvios.frx":04E0
         FormatStyle(6)  =   "frmReqVsRealAvios.frx":0598
         ImageCount      =   0
         PrinterProperties=   "frmReqVsRealAvios.frx":0678
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1020
      Left            =   120
      TabIndex        =   6
      Top             =   -15
      Width           =   10245
      Begin VB.OptionButton OpTelas 
         Caption         =   "Telas"
         Height          =   255
         Left            =   6840
         TabIndex        =   13
         Top             =   700
         Width           =   735
      End
      Begin VB.OptionButton OpAvios 
         Caption         =   "Avios"
         Height          =   255
         Left            =   5640
         TabIndex        =   12
         Top             =   700
         Value           =   -1  'True
         Width           =   855
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   495
         Left            =   8760
         TabIndex        =   4
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~True~True~&Buscar~0~0~1~~0~False~False~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.TextBox TxtEstilo 
         Height          =   300
         Left            =   5340
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   300
         Width           =   3030
      End
      Begin VB.TextBox TxtOP 
         Height          =   285
         Left            =   4575
         MaxLength       =   5
         TabIndex        =   3
         Top             =   300
         Width           =   735
      End
      Begin VB.CommandButton cmdBuscaFabrica 
         Caption         =   "..."
         Height          =   330
         Left            =   1440
         TabIndex        =   1
         Top             =   285
         Width           =   330
      End
      Begin VB.TextBox txtNom_Fabrica 
         Height          =   285
         Left            =   1755
         TabIndex        =   2
         Top             =   300
         Width           =   1800
      End
      Begin VB.TextBox txtAbr_Fabrica 
         Height          =   285
         Left            =   810
         MaxLength       =   5
         TabIndex        =   0
         Top             =   300
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "OP"
         Height          =   240
         Left            =   4230
         TabIndex        =   9
         Top             =   345
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fabrica"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   330
         Width           =   525
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   7200
      Top             =   5520
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmReqVsRealAvios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo As String
Public Descripcion As String

Dim Strsql As String

Private Sub cmdBuscaFabrica_Click()
    Call Me.BUSCA_FABRICA(3)
End Sub

Private Sub CmdImprimir_Click()
    CARGA_GRID
    If Me.gexList.RowCount = 0 Then
        MsgBox "No hay datos para imprimir", vbInformation, Me.Caption
        Exit Sub
    End If
    IMPRESION
End Sub

Public Sub IMPRESION()
Dim varTipItem As String
On Error GoTo ErrorImpresion
    Dim oo As Object
    
        If OpAvios.Value = True Then
            varTipItem = "I"
        Else
            varTipItem = "T"
        End If
        
        Strsql = "select ruta_logo from seguridad..seg_empresas where cod_Empresa='" & vemp1 & "'"
        
        Set oo = CreateObject("excel.application")
        oo.Workbooks.Open vRuta & "\ReqRealAvios-Telas.xlt"
        oo.Visible = True
        
        oo.Run "REPORTE", Me.gexList.ADORecordset, varTipItem, cConnect, DevuelveCampo(Strsql, cConnect)
        Screen.MousePointer = vbNormal
        oo.Visible = True
        Set oo = Nothing


    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte  " & Err.Description, vbCritical, "Impresion"

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Strsql = "SELECT Abr_Fabrica FROM TG_FABRICA"
    Me.txtAbr_Fabrica.Text = DevuelveCampo(Strsql, cConnect)
    If Trim(Me.txtAbr_Fabrica.Text) <> "" Then
        Strsql = "SELECT Nom_Fabrica as 'Descripción' FROM TG_FABRICA WHERE Abr_Fabrica = '" & Trim(Me.txtAbr_Fabrica.Text) & "'"
        Me.txtNom_Fabrica.Text = Trim(DevuelveCampo(Strsql, cConnect))
        'TxtOP.SetFocus
    End If

    Strsql = "select tipo_orden from tg_control"
    Label1.Caption = DevuelveCampo(Strsql, cConnect)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    If Trim(txtAbr_Fabrica.Text) = "" Then
        MsgBox "Seleccione la Fabrica", vbInformation, Me.Caption
        txtAbr_Fabrica.SetFocus
    End If
    
    If Trim(TxtOP.Text) = "" Then
        MsgBox "Ingrese la " & Label1.Caption, vbInformation, Me.Caption
        TxtOP.SetFocus
    End If
    
    CARGA_GRID
End Sub

Sub CARGA_GRID()
On Error GoTo hand
Dim sCod_Fabrica As String

Strsql = "select cod_fabrica from tg_fabrica where abr_fabrica='" & Me.txtAbr_Fabrica.Text & "'"
sCod_Fabrica = DevuelveCampo(Strsql, cConnect)

If OpTelas.Value = True Then
    Strsql = "EXEC SM_AVANCES_TELA_TENIDA_ORDEN '" & sCod_Fabrica & "','" & Me.TxtOP.Text & "','T','T'"
Else
    Strsql = "EXEC SM_MUESTRAS_CONSUMOS_REQ_ORDEN '" & sCod_Fabrica & "','" & Me.TxtOP.Text & "'"
End If
                                
VB.Screen.MousePointer = 11
Set Me.gexList.ADORecordset = CargarRecordSetDesconectado(Strsql, cConnect)
VB.Screen.MousePointer = 0

If OpAvios.Value = True Then ConfigurarGrid


Exit Sub
hand:
ErrorHandler Err, "CARGA_GRID"
End Sub

Sub ConfigurarGrid()

    gexList.Columns("Cod.Avio").Width = 1200
    gexList.Columns("Descripcion").Width = 2500
    gexList.Columns("UN").Width = 700
    gexList.Columns("Origen").Width = 700
    gexList.Columns("Requerida").Width = 1000
    gexList.Columns("Comprada").Width = 1000
    gexList.Columns("Recibida").Width = 1000
    
End Sub

Private Sub txtAbr_Fabrica_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtAbr_Fabrica.Text) = "" Then
            Call Me.BUSCA_FABRICA(3)
        Else
            Call Me.BUSCA_FABRICA(1)
        End If
    End If
End Sub

Public Sub BUSCA_FABRICA(Tipo As Integer)
    Select Case Tipo
        Case 1:
                    Strsql = "SELECT Nom_Fabrica as 'Descripción' FROM TG_FABRICA WHERE Abr_Fabrica = '" & Trim(Me.txtAbr_Fabrica.Text) & "' ORDER BY Abr_Fabrica"
                    Me.txtNom_Fabrica.Text = Trim(DevuelveCampo(Strsql, cConnect))
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim Rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.sQuery = "SELECT Abr_Fabrica as 'Código', Nom_Fabrica as 'Descripción' FROM TG_FABRICA WHERE Nom_Fabrica LIKE '%" & Trim(Me.txtNom_Fabrica.Text) & "%' ORDER BY Abr_Fabrica"
                    Else
                        oTipo.sQuery = "SELECT Abr_Fabrica as 'Código', Nom_Fabrica as 'Descripción' FROM TG_FABRICA ORDER BY Abr_Fabrica"
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.Show 1
                    If Codigo <> "" Then
                        Me.txtAbr_Fabrica.Text = Trim(Codigo)
                        Me.txtNom_Fabrica.Text = Trim(Descripcion)
                    End If
                    Set oTipo = Nothing
                    Set Rs = Nothing
    End Select
    Codigo = "": Descripcion = ""
    Me.TxtOP.SetFocus
End Sub

Private Sub txtNom_Fabrica_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Me.BUSCA_FABRICA(2)
    End If
End Sub

Private Sub TxtOP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Dim sCod_Fabrica As String

        Strsql = "select cod_fabrica from tg_fabrica where abr_fabrica='" & Me.txtAbr_Fabrica.Text & "'"
        sCod_Fabrica = DevuelveCampo(Strsql, cConnect)

        TxtOP.Text = Format(Trim(TxtOP.Text), "00000")
        If DevuelveCampo("select count(*) from es_Ordpro where cod_fabrica='" & sCod_Fabrica & "' AND cod_ordpro = '" & TxtOP.Text & "'", cConnect) > 0 Then
            Strsql = "SELECT cod_fabrica FROM TG_FABRICA WHERE Abr_Fabrica = '" & Trim(Me.txtAbr_Fabrica.Text) & "'"
            Me.TxtEstilo.Text = DevuelveCampo("SELECT b.Des_EstPro FROM   ES_OrdPro  a , ES_EstPRo b WHERE  a.Cod_EstPro = b.Cod_EstPRo AND a.Cod_Fabrica= '" & DevuelveCampo(Strsql, cConnect) & "' AND a.Cod_OrdPro = '" & TxtOP.Text & "'", cConnect)
            Me.FunctButt1.SetFocus
        Else
            MsgBox "Codigo de " & Label1.Caption & " no existe", vbInformation, Me.Caption
        End If
    End If
End Sub
