VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmFlujoProduccionDiaria 
   Caption         =   "Flujo Produccion Diaria"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   510
      Left            =   8895
      TabIndex        =   12
      Top             =   6360
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmFujoProduccionDiaria.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame2 
      Height          =   5295
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   11295
      Begin GridEX20.GridEX gexList 
         Height          =   5055
         Left            =   120
         TabIndex        =   10
         Top             =   150
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   8916
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         CursorLocation  =   3
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         ColumnHeaderHeight=   285
         FmtConditionsCount=   1
         FmtCondition(1) =   "frmFujoProduccionDiaria.frx":0094
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmFujoProduccionDiaria.frx":0158
         FormatStyle(2)  =   "frmFujoProduccionDiaria.frx":0290
         FormatStyle(3)  =   "frmFujoProduccionDiaria.frx":0340
         FormatStyle(4)  =   "frmFujoProduccionDiaria.frx":03F4
         FormatStyle(5)  =   "frmFujoProduccionDiaria.frx":04CC
         FormatStyle(6)  =   "frmFujoProduccionDiaria.frx":0584
         ImageCount      =   0
         PrinterProperties=   "frmFujoProduccionDiaria.frx":0664
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   495
         Left            =   9960
         TabIndex        =   11
         Top             =   120
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   8160
         TabIndex        =   6
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   65273857
         CurrentDate     =   39027
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   6000
         TabIndex        =   5
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   65273857
         CurrentDate     =   37750
      End
      Begin VB.TextBox txtAbr_Fabrica 
         Height          =   285
         Left            =   960
         MaxLength       =   5
         TabIndex        =   4
         Top             =   255
         Width           =   630
      End
      Begin VB.TextBox txtNom_Fabrica 
         Height          =   285
         Left            =   1905
         TabIndex        =   3
         Top             =   255
         Width           =   1800
      End
      Begin VB.CommandButton cmdBuscaFabrica 
         Caption         =   "..."
         Height          =   330
         Left            =   1590
         TabIndex        =   2
         Top             =   240
         Width           =   330
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta :"
         Height          =   255
         Left            =   7560
         TabIndex        =   8
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Desde :"
         Height          =   255
         Left            =   5280
         TabIndex        =   7
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Fabrica :"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   280
         Width           =   615
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   240
      Top             =   6360
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmFlujoProduccionDiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo As String
Public Descripcion As String
Public strSQL As String

Private Sub cmdBuscaFabrica_Click()
    Call Me.BUSCA_FABRICA(3)
End Sub

Private Sub DTPicker1_Click()

    'Me.DTPicker2.MinDate = Me.DTPicker1.Value
    'Me.DTPicker2.Value = Me.DTPicker1.Value
    'Me.DTPicker2.MaxDate = Me.DTPicker1.Value + 45
    
    
    
End Sub


Private Sub Form_Load()
    strSQL = "SELECT Abr_Fabrica FROM TG_FABRICA"
    Me.txtAbr_Fabrica.Text = DevuelveCampo(strSQL, cConnect)
    If Trim(Me.txtAbr_Fabrica.Text) <> "" Then
        strSQL = "SELECT Nom_Fabrica as 'Descripción' FROM TG_FABRICA WHERE Abr_Fabrica = '" & Trim(Me.txtAbr_Fabrica.Text) & "'"
        Me.txtNom_Fabrica.Text = Trim(DevuelveCampo(strSQL, cConnect))
    End If
    Me.DTPicker1.Value = Date
    'Me.DTPicker2.MinDate = Me.DTPicker1.Value
    'Me.DTPicker2.MaxDate = Date + 45
    
    Me.DTPicker2.Value = Date
    Me.DTPicker1.Value = DateAdd("d", -7, Date)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    If Me.DTPicker1.Value > Me.DTPicker2.Value Then
        MsgBox "La fecha inicial no puede mayor que la final, verifique", vbInformation, Me.Caption
        Me.DTPicker2.SetFocus
        Exit Sub
    End If
    
    carga_grid
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "IMPRIMIR"
            Call Reporte
        Case "SALIR"
            Unload Me
    End Select
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
                    strSQL = "SELECT Nom_Fabrica as 'Descripción' FROM TG_FABRICA WHERE Abr_Fabrica = '" & Trim(Me.txtAbr_Fabrica.Text) & "' ORDER BY Abr_Fabrica"
                    Me.txtNom_Fabrica.Text = Trim(DevuelveCampo(strSQL, cConnect))
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim rs As New ADODB.Recordset
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
                    Set rs = Nothing
    End Select
    Codigo = "": Descripcion = ""
    Me.DTPicker1.SetFocus
End Sub

Private Sub txtNom_Fabrica_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Me.BUSCA_FABRICA(2)
    End If
End Sub

Sub carga_grid()
On Error GoTo hand
Dim sCod_Fabrica As String

strSQL = "select cod_fabrica from tg_fabrica where abr_fabrica='" & Me.txtAbr_Fabrica.Text & "'"
sCod_Fabrica = DevuelveCampo(strSQL, cConnect)

strSQL = "EXEC sm_acumula_produccion_diaria_corte_costura_globales '" & sCod_Fabrica & "','" & Me.DTPicker1.Value & "','" & Me.DTPicker2.Value & "'"
                                
VB.Screen.MousePointer = 11
Set Me.gexList.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
ConfigurarGrid
VB.Screen.MousePointer = 0


Exit Sub
hand:
ErrorHandler err, "CARGA_GRID"
End Sub

Sub ConfigurarGrid()
Dim fmtCon As JSFmtCondition
Dim i As Integer

With gexList.FmtConditions
        .ApplyGroupCondition = True
        .ShowGroupConditionCount = True
        .GroupConditionCountTitle = "Tipo_dato"
        Set fmtCon = .GroupCondition
End With
Set fmtCon = gexList.FmtConditions.Add(1, jgexContains, "TOTAL")

    'fmtCon.FormatStyle.BackColor = &HC0FFFF
    fmtCon.SetCondition 1, jgexContains, "TOTAL"
    fmtCon.FormatStyle.ForeColor = &H80000002
    fmtCon.FormatStyle.FontBold = True
End Sub

Public Sub Reporte()
On Error GoTo ErrorImpresion
    Dim oo As Object
    strSQL = "select ruta_logo from seguridad..seg_empresas where cod_Empresa='" & vemp1 & "'"
    
    Set oo = CreateObject("excel.application")
    oo.workbooks.Open vRuta & "\FlujoProducDiaria.xlt"
    oo.Visible = True
    
    oo.run "REPORTE", Me.gexList.ADORecordset, DevuelveCampo(strSQL, cConnect)
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte  " & err.Description, vbCritical, "Impresion"
End Sub

