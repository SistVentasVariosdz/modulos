VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Registro_Ventas 
   Caption         =   "Registro De Ventas"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13605
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   13605
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra_imprimir 
      Height          =   1935
      Left            =   5520
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   600
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1440
         TabIndex        =   9
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   70909953
         CurrentDate     =   41000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   720
         Width           =   450
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   630
      Left            =   120
      TabIndex        =   3
      Top             =   6480
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   1111
      Custom          =   $"Frm_Registro_Ventas.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1500
      ControlHeigth   =   600
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX2 
      Height          =   5535
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   9763
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   2
      Column(1)       =   "Frm_Registro_Ventas.frx":0105
      Column(2)       =   "Frm_Registro_Ventas.frx":01CD
      FormatStylesCount=   6
      FormatStyle(1)  =   "Frm_Registro_Ventas.frx":0271
      FormatStyle(2)  =   "Frm_Registro_Ventas.frx":03A9
      FormatStyle(3)  =   "Frm_Registro_Ventas.frx":0459
      FormatStyle(4)  =   "Frm_Registro_Ventas.frx":050D
      FormatStyle(5)  =   "Frm_Registro_Ventas.frx":05E5
      FormatStyle(6)  =   "Frm_Registro_Ventas.frx":069D
      ImageCount      =   0
      PrinterProperties=   "Frm_Registro_Ventas.frx":077D
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13575
      Begin VB.OptionButton Option3 
         Caption         =   "Todos"
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Exportacion"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nacional"
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   120
         Value           =   -1  'True
         Width           =   1095
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   495
         Left            =   12120
         TabIndex        =   5
         Top             =   360
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
      Begin GridEX20.GridEX GridEX1 
         Height          =   30
         Left            =   0
         TabIndex        =   1
         Top             =   1200
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   53
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ScrollToolTipColumn=   ""
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "Frm_Registro_Ventas.frx":0955
         Column(2)       =   "Frm_Registro_Ventas.frx":0A1D
         FormatStylesCount=   6
         FormatStyle(1)  =   "Frm_Registro_Ventas.frx":0AC1
         FormatStyle(2)  =   "Frm_Registro_Ventas.frx":0BF9
         FormatStyle(3)  =   "Frm_Registro_Ventas.frx":0CA9
         FormatStyle(4)  =   "Frm_Registro_Ventas.frx":0D5D
         FormatStyle(5)  =   "Frm_Registro_Ventas.frx":0E35
         FormatStyle(6)  =   "Frm_Registro_Ventas.frx":0EED
         ImageCount      =   0
         PrinterProperties=   "Frm_Registro_Ventas.frx":0FCD
      End
      Begin MSComCtl2.DTPicker dtpAnoMes 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM yyyy"
         Format          =   70909955
         CurrentDate     =   37887
      End
      Begin VB.Label Label3 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5760
         TabIndex        =   15
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año/Mes :"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   435
         Width           =   750
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   1200
      Top             =   5640
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "Frm_Registro_Ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cadena As String
Public origen As String
Public SREPORTE As String, strSQL As String
Dim mes As String
Dim flag As Boolean

Private Sub Command1_Click()
SREPORTE = ""
If vemp = "01" Or vemp = "03" Or vemp = "07" Or vemp = "09" Then
    SREPORTE = Trim(DevuelveCampo("SELECT RptHialpesaRegVentas FROM CN_Control ", cCONNECT))
Else
    SREPORTE = Trim(DevuelveCampo("SELECT RptCottonRegVentas FROM CN_Control ", cCONNECT))
End If
    Reporte
End Sub

Private Sub Command2_Click()
fra_imprimir.Visible = False
End Sub

Private Sub dtpAnoMes_Change()
Me.Label3 = DevuelveCampo("select DBO.CN_Devuelve_Estado_Datos_LDP_DDP  ('" & Year(dtpAnoMes) & "','" & Format(Month(dtpAnoMes), "00") & "')", cCONNECT)
End Sub

Private Sub Form_Load()
Dim sSeguridad  As String

  sSeguridad = get_botones1(Me, vper, vemp, Me.Name)
  Me.FunctButt1.FunctionsUser = sSeguridad
  
dtpAnoMes = Date
DTPicker1 = Date
origen = "N"
Me.Label3 = DevuelveCampo("select DBO.CN_Devuelve_Estado_Datos_LDP_DDP  ('" & Year(dtpAnoMes) & "','" & Month(dtpAnoMes) & "')", cCONNECT)
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim Msg As Variant
Select Case ActionName
    Case "IMPRIMIR"
        fra_imprimir.Visible = True
    Case "NC"
    Msg = MsgBox("¿Esta seguro de Verificar Asociacion Factura - NC?", vbYesNo)
        If Msg = vbNo Then Exit Sub
        Call Actualiza_NC
    Case "SALIR"
        Unload Me
        
End Select
End Sub

Private Sub Actualiza_NC()
    On Error GoTo errorx
    Dim sSql As String
    Dim aMess(4), i As Integer
    
      
    ExecuteCommandSQL cCONNECT, "HIL_VENTAS_GENERA_DOCUM_AUTORIZADOS '" & Year(dtpAnoMes) & "','" & Format(Month(dtpAnoMes), "00") & "'"
    Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
     
    Exit Sub
    Resume
errorx:
        errores err.Number
End Sub

Sub CARGA_GRID()
Dim vopcion As String
Dim anio As String

flag = False
anio = Year(dtpAnoMes.Value)
mes = Format(Month(dtpAnoMes), "00")

If anio & mes >= DevuelveCampo("SELECT Ano_Periodo_Nuevo_Formato  FROM CN_CONTROL", cCONNECT) Then
    strSQL = "exec Ventas_Emision_Registro_Mensual_SUNAT '" & anio & "','" & mes & "','" & origen & "'"
    flag = True
Else
    strSQL = "exec Ventas_Emision_Registro_Mensual '" & anio & "','" & mes & "','" & origen & "'"

End If

cadena = strSQL


Set GridEX2.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

If flag = True Then
GridEX2.Columns("clase").Visible = False
GridEX2.Columns("orden").Visible = False
GridEX2.Columns("num_registro").Visible = False
GridEX2.FrozenColumns = 7
Else
GridEX2.Columns("Doc_Sunat").Width = 900
GridEX2.Columns("Doc").Width = 1170
GridEX2.Columns("fECHA").Width = 1095
GridEX2.Columns("CLIENTE").Width = 3480
GridEX2.FrozenColumns = 4

End If

End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "BUSCAR"
        CARGA_GRID
End Select
End Sub

Sub Reporte()
On Error GoTo ErrorImpresion
Dim smes As String, sRutaLogo As String
Dim oo As Object, lvSql As String

    strSQL = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
    sRutaLogo = DevuelveCampo(strSQL, cCONNECT)
    
    smes = DevuelveCampo("select dbo.uf_nombre_mes('" & Format(CInt(Format(Month(DateAdd("m", -0, dtpAnoMes)), "00")), "0#") & "',1)", cCONNECT)
    Set oo = CreateObject("excel.application")

    If flag = True Then
            oo.Workbooks.Open vRuta & "\" & "Rpt_Registro_Ventas_NuevoFormato.xlt"
    Else
            oo.Workbooks.Open vRuta & "\" & SREPORTE
    End If

    
    
    oo.Visible = True
    oo.DisplayAlerts = False
    If SREPORTE = "Rpt_Registro_Ventas.XLT" Then
        oo.Run "reporte", Year(dtpAnoMes), smes, DTPicker1, cadena, origen, cCONNECT, sRutaLogo, mes
    Else
        oo.Run "reporte", Year(dtpAnoMes), smes, DTPicker1, cadena, origen, cCONNECT
    End If
    Set oo = Nothing
    Exit Sub
    'DevuelveCampo("select dbo.uf_nombre_mes('" & Format(CInt(Format(Month(DateAdd("m", -0, dtpAnoMes)), "00")), "0#") & "',1)", cCONNECT) & "-" & Year(DateAdd("m", -0, dtpAnoMes))
    
    
    
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte de Guia de Remisión " & err.Description, vbCritical, "Impresion"
End Sub


Private Sub Option1_Click()
origen = "N"
End Sub

Private Sub Option2_Click()
origen = "E"
End Sub

Private Sub Option3_Click()
origen = ""
End Sub
