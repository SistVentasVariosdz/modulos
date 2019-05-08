VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form frmControlProdMensual 
   Caption         =   "Control de Produccion Mensual"
   ClientHeight    =   6435
   ClientLeft      =   1860
   ClientTop       =   1815
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   10590
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   510
      Left            =   7920
      TabIndex        =   8
      Top             =   5640
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmControlProdMensual.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame2 
      Height          =   4575
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   10335
      Begin GridEX20.GridEX gexList 
         Height          =   4335
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7646
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmControlProdMensual.frx":0094
         Column(2)       =   "frmControlProdMensual.frx":015C
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmControlProdMensual.frx":0200
         FormatStyle(2)  =   "frmControlProdMensual.frx":0338
         FormatStyle(3)  =   "frmControlProdMensual.frx":03E8
         FormatStyle(4)  =   "frmControlProdMensual.frx":049C
         FormatStyle(5)  =   "frmControlProdMensual.frx":0574
         FormatStyle(6)  =   "frmControlProdMensual.frx":062C
         ImageCount      =   0
         PrinterProperties=   "frmControlProdMensual.frx":070C
      End
   End
   Begin VB.Frame Frame1 
      Height          =   840
      Left            =   120
      TabIndex        =   14
      Top             =   60
      Width           =   10335
      Begin VB.TextBox txtAnioF 
         Height          =   285
         Left            =   5310
         TabIndex        =   5
         Top             =   495
         Width           =   555
      End
      Begin VB.TextBox txtMesF 
         Height          =   285
         Left            =   6285
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Resumido"
         Height          =   255
         Left            =   7440
         TabIndex        =   11
         Top             =   510
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Detallado"
         Height          =   255
         Left            =   7440
         TabIndex        =   10
         Top             =   255
         Value           =   -1  'True
         Width           =   1095
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   495
         Left            =   9000
         TabIndex        =   7
         Top             =   195
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
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   6661
         TabIndex        =   12
         Top             =   165
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "TxtMes"
         BuddyDispid     =   196615
         OrigLeft        =   7440
         OrigTop         =   240
         OrigRight       =   7680
         OrigBottom      =   495
         Max             =   12
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox TxtMes 
         Height          =   285
         Left            =   6285
         TabIndex        =   4
         Top             =   165
         Width           =   615
      End
      Begin VB.TextBox txtAnio 
         Height          =   285
         Left            =   5310
         TabIndex        =   3
         Top             =   180
         Width           =   555
      End
      Begin VB.CommandButton cmdBuscaFabrica 
         Caption         =   "..."
         Height          =   330
         Left            =   1470
         TabIndex        =   1
         Top             =   315
         Width           =   330
      End
      Begin VB.TextBox txtNom_Fabrica 
         Height          =   285
         Left            =   1785
         TabIndex        =   2
         Top             =   330
         Width           =   1800
      End
      Begin VB.TextBox txtAbr_Fabrica 
         Height          =   285
         Left            =   840
         MaxLength       =   5
         TabIndex        =   0
         Top             =   330
         Width           =   630
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   6660
         TabIndex        =   13
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMesF"
         BuddyDispid     =   196612
         OrigLeft        =   7440
         OrigTop         =   240
         OrigRight       =   7680
         OrigBottom      =   495
         Max             =   12
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label7 
         Caption         =   "Hasta :"
         Height          =   255
         Left            =   4185
         TabIndex        =   22
         Top             =   510
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Desde :"
         Height          =   255
         Left            =   4170
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Left            =   4935
         TabIndex        =   20
         Top             =   555
         Width           =   330
      End
      Begin VB.Label Label4 
         Caption         =   "Mes:"
         Height          =   255
         Left            =   5910
         TabIndex        =   19
         Top             =   540
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Mes:"
         Height          =   255
         Left            =   5910
         TabIndex        =   18
         Top             =   225
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Left            =   4935
         TabIndex        =   17
         Top             =   240
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "Fabrica :"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   240
      Top             =   5640
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmControlProdMensual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo As String
Public Descripcion As String
Dim strSQL As String
Dim sCod_Fabrica As String

Private Sub cmdBuscaFabrica_Click()
    Call Me.BUSCA_FABRICA(3)
End Sub

Private Sub Form_Load()
    strSQL = "SELECT Abr_Fabrica FROM TG_FABRICA"
    Me.txtAbr_Fabrica.Text = DevuelveCampo(strSQL, cConnect)
    If Trim(Me.txtAbr_Fabrica.Text) <> "" Then
        strSQL = "SELECT Nom_Fabrica as 'Descripción' FROM TG_FABRICA WHERE Abr_Fabrica = '" & Trim(Me.txtAbr_Fabrica.Text) & "'"
        Me.txtNom_Fabrica.Text = Trim(DevuelveCampo(strSQL, cConnect))
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Call CARGA_GRID
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "IMPRIMIR"
            If gexList.RowCount = 0 Then
                MsgBox "No hoy datos a imprimir", vbInformation, Me.Caption
                Exit Sub
            Else
                Reporte
            End If
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
    'Me.DTPicker1.SetFocus
End Sub

Private Sub TxtMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.TxtMes.Text) <> "" Then Me.FunctButt1.SetFocus
    Else
        Call SoloNumeros(Me.TxtMes, KeyAscii, False, 0, 2)
    End If
End Sub

Private Sub TxtMes_LostFocus()
    If Trim(TxtMes.Text) > 12 Or Trim(TxtMes.Text) < 1 Then
        MsgBox "Mes incorrecto", vbCritical, "Mensaje"
        TxtMes.SetFocus
    End If
End Sub

Private Sub txtNom_Fabrica_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Me.BUSCA_FABRICA(2)
    End If
End Sub
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.txtAnio.Text) <> "" Then Me.TxtMes.SetFocus
    Else
        Call SoloNumeros(Me.txtAnio, KeyAscii, False, 0, 4)
    End If
End Sub

Sub CARGA_GRID()
On Error GoTo hand
Dim dtIni As Date, dtFin As Date

dtIni = "01 " & MonthName(TxtMes) & " " & txtAnio
dtFin = "01 " & MonthName(txtMesF) & " " & txtAnioF

If dtFin < dtIni Then
    MsgBox "La fecha final debe ser mayor o igual a la fecha inicial", vbOKOnly _
    + vbInformation, "Reporte"
    txtAnio.SetFocus
    Exit Sub
End If

VB.Screen.MousePointer = 11
strSQL = "select cod_fabrica from tg_fabrica where abr_fabrica='" & Me.txtAbr_Fabrica.Text & "'"
sCod_Fabrica = DevuelveCampo(strSQL, cConnect)

If Option1.Value Then
    strSQL = "EXEC sm_avances_confecciones_orden_ano_mes_COLOR '" & sCod_Fabrica & "','" & _
             Me.txtAnio.Text & "','" & Right("0" & Trim(Me.TxtMes.Text), 2) & "', '" & _
             Me.txtAnioF.Text & "','" & Right("0" & Trim(Me.txtMesF.Text), 2) & "'"
Else
    strSQL = "EXEC sm_avances_confecciones_orden_ano_mes '" & sCod_Fabrica & "', '" & _
             Me.txtAnio.Text & "', '" & Right("0" & Trim(Me.TxtMes.Text), 2) & "', '" & _
             Me.txtAnioF.Text & "', '" & Right("0" & Trim(Me.txtMesF.Text), 2) & "'"
End If
                                
Set Me.gexList.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
VB.Screen.MousePointer = 0
'ConfigurarGrid


Exit Sub
hand:
ErrorHandler err, "CARGA_GRID"
End Sub

Sub Reporte()
On Error GoTo ErrorImpresion

Dim dtIni As Date, dtFin As Date
    
    dtIni = "01 " & MonthName(TxtMes) & " " & txtAnio
    dtFin = "01 " & MonthName(txtMesF) & " " & txtAnioF
    
    If dtFin < dtIni Then
        MsgBox "La fecha final debe ser mayor o igual a la fecha inicial", vbOKOnly _
        + vbInformation, "Reporte"
        txtAnio.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    Dim oo As Object
    strSQL = "select ruta_logo from seguridad..seg_empresas where cod_Empresa='" & vemp1 & "'"
    
    Set oo = CreateObject("excel.application")
    If Option1.Value Then
        oo.Workbooks.Open vRuta & "\FlujoProdMensualD.xlt"
    Else
        oo.Workbooks.Open vRuta & "\FlujoProdMensualR.xlt"
    End If
    oo.Visible = True
    
    oo.Run "REPORTE", sCod_Fabrica, Trim(Me.txtAnio.Text), Right("0" & Trim(Me.TxtMes.Text), 2), _
    Trim(Me.txtAnioF.Text), Right("0" & Trim(Me.txtMesF.Text), 2), cConnect, DevuelveCampo(strSQL, cConnect)
    
    'oo.Run "REPORTE", "2002", "01", gexLista.ADORecordset, "\\SERVER02\LOGOS\LOGO.BMP"
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrorImpresion:
    Screen.MousePointer = 0
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte" & err.Description, vbCritical, "Impresion"
End Sub
