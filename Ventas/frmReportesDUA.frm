VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReportesDUA 
   Caption         =   "Reportes DUA"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmCerrarDua 
      Caption         =   "Cerrar Dua"
      Height          =   1815
      Left            =   1680
      TabIndex        =   13
      Top             =   1320
      Width           =   3375
      Begin VB.TextBox txtMes 
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtAno 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   480
         Left            =   720
         TabIndex        =   18
         Top             =   1200
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   847
         Custom          =   $"frmReportesDUA.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1000
         ControlHeigth   =   450
         ControlSeparator=   110
      End
      Begin VB.Label Label6 
         Caption         =   "Mes  :"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Ano :"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CheckBox chkDetallado 
      Alignment       =   1  'Right Justify
      Caption         =   "Detallado"
      Height          =   285
      Left            =   4320
      TabIndex        =   8
      Top             =   1395
      Width           =   1665
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   1935
      TabIndex        =   3
      Top             =   825
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "yyyy/MM"
      Format          =   56557571
      CurrentDate     =   39018
   End
   Begin VB.OptionButton optDUASPendientes 
      Caption         =   "Reporte de DUAS Pendientes"
      Height          =   315
      Left            =   630
      TabIndex        =   1
      Top             =   2955
      Width           =   2535
   End
   Begin VB.OptionButton optRepAnoMes 
      Caption         =   "Relación de DUAS por Año/Mes Embarque"
      Height          =   315
      Left            =   645
      TabIndex        =   0
      Top             =   345
      Value           =   -1  'True
      Width           =   3390
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   570
      Left            =   1320
      TabIndex        =   2
      Top             =   3480
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   1005
      Custom          =   $"frmReportesDUA.frx":0093
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1200
      ControlHeigth   =   550
      ControlSeparator=   110
   End
   Begin MSComCtl2.DTPicker dtpFec_Numerac_Ini 
      Height          =   285
      Left            =   4740
      TabIndex        =   6
      Top             =   945
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   503
      _Version        =   393216
      Format          =   56557569
      CurrentDate     =   367
   End
   Begin MSComCtl2.DTPicker dtpFec_Numerac_Fin 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   1380
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   503
      _Version        =   393216
      Format          =   56557569
      CurrentDate     =   39018
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   285
      Left            =   4680
      TabIndex        =   10
      Top             =   1965
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   503
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "yyyy/MM"
      Format          =   56557571
      CurrentDate     =   39191
   End
   Begin MSComCtl2.DTPicker dtpFec_Numerac_Anterior 
      Height          =   285
      Left            =   4680
      TabIndex        =   12
      Top             =   2400
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   56557569
      CurrentDate     =   39147
   End
   Begin VB.Label Label4 
      Caption         =   "fecha de numeracion sea mayor a"
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Incluir facts embarcadas en periodo"
      Height          =   255
      Left            =   930
      TabIndex        =   9
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Límite de Numeración"
      Height          =   585
      Left            =   930
      TabIndex        =   5
      Top             =   1275
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "Periodo Embarque"
      Height          =   450
      Left            =   930
      TabIndex        =   4
      Top             =   765
      Width           =   780
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   4680
      Top             =   3480
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmReportesDUA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String


Private Sub Form_Load()
    DTPicker1.Value = Date - 30
    dtpFec_Numerac_Fin.Value = Date - 30
        
    Me.DTPicker2.Value = DateAdd("m", -1, Me.DTPicker1.Value)
    Me.DTPicker2.Value = ""
    
    Me.dtpFec_Numerac_Anterior.Value = DateAdd("d", -10, Me.dtpFec_Numerac_Fin)
    Me.dtpFec_Numerac_Anterior.Value = ""
    frmCerrarDua.Visible = False
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "IMPRIMIR"
            If optRepAnoMes Then
                If chkDetallado.Value = "1" Then
                    Imprimir 3
                Else
                    Imprimir 4
                End If
            Else
                Imprimir 2
            End If
        Case "CERRARDUA"
            Me.frmCerrarDua.Visible = True
            Me.txtAno = Year(Date)
            Me.txtMes = Format(Month(Date), "00")
            Me.txtAno.SetFocus
            SelectionText Me.txtAno
        Case "SALIR"
            Unload Me
    End Select
End Sub

Sub Imprimir(opcion As Integer)
On Error GoTo hand
Dim oo As Object
Dim strSQL  As String, titulo As String, fecha As String, _
    fecOpcion As String, Ruta As String, iResp As Integer, _
    sFec_Numeracion_Ini As String, sFec_Numeracion_Fin As String
    
    If opcion = 1 Then
        strSQL = "CN_MUESTRA_LISTADOS_DUAS '1','" & DTPicker1.Year & "','" & Format(DTPicker1.Month, "00") & "'"
        titulo = optRepAnoMes.Caption
        fecha = DTPicker1.Year & "/" & Format(DTPicker1.Month, "00")
        fecOpcion = "Fecha"
    End If
    
    If opcion = 2 Then
        strSQL = "CN_MUESTRA_LISTADOS_DUAS '2','" & DTPicker1.Year & "','" & Format(DTPicker1.Month, "00") & "'"
        titulo = optDUASPendientes.Caption
        fecha = " "
        fecOpcion = " "
    End If
    
    If opcion = 3 Then
        strSQL = "CN_MUESTRA_LISTADOS_DUAS '3','" & DTPicker1.Year & "','" & Format(DTPicker1.Month, "00") & "','" & dtpFec_Numerac_Ini.Value & "','" & dtpFec_Numerac_Fin.Value & "','" & Me.DTPicker2.Year & "','" & Format(Me.DTPicker2.Month, "00") & "','" & Me.dtpFec_Numerac_Anterior & "'"
        titulo = optRepAnoMes.Caption & " - Detallado"
        fecha = " "
        fecOpcion = " "
    End If
    
    If opcion = 4 Then
        strSQL = "CN_MUESTRA_LISTADOS_DUAS '4','" & DTPicker1.Year & "','" & Format(DTPicker1.Month, "00") & "','" & dtpFec_Numerac_Ini.Value & "','" & dtpFec_Numerac_Fin.Value & "','" & Me.DTPicker2.Year & "','" & Format(Me.DTPicker2.Month, "00") & "','" & Me.dtpFec_Numerac_Anterior & "'"
        titulo = optRepAnoMes.Caption
        fecha = " "
        fecOpcion = " "
    End If
    
    iResp = MsgBox("Imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir")
    
    Set oo = CreateObject("excel.application")
    If opcion <= 3 Then
        Ruta = vRuta & "\RptRepDUA." & IIf((iResp = vbYes), "XLT", "OTS")
    End If
    
    If opcion = 4 Then
        Ruta = vRuta & "\RptRepDUAResumen." & IIf((iResp = vbYes), "XLT", "OTS")
    End If
    
    If iResp = vbYes Then
        oo.Workbooks.Open Ruta
        oo.Visible = True
        oo.DisplayAlerts = False
        
        oo.Run "ReporteDuas", cCONNECT, strSQL, fecha, titulo, fecOpcion, dtpFec_Numerac_Ini.Value, dtpFec_Numerac_Fin.Value, DTPicker1.Year & "/" & Format(DTPicker1.Month, "00")
    Else
        Set oo = CreateObject("ooBusiness.Calc")
        oo.OfficeTemplateSheet = Ruta
        oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
        oo.MacroLibraryName = "Library1"
        oo.MacroModuleName = "Module1"
        oo.MacroName = "ReporteDuas"
        
        oo.Run cCONNECT, strSQL, fecha, titulo, fecOpcion, dtpFec_Numerac_Ini.Value, dtpFec_Numerac_Fin.Value, DTPicker1.Year & "/" & Format(DTPicker1.Month, "00")
    End If
    Set oo = Nothing
    Unload Me
Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub


Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "CERRAR"
        SALVAR_DATOS
        Me.frmCerrarDua.Visible = False
    Case "CANCELAR"
        Me.frmCerrarDua.Visible = False
End Select
End Sub

Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr

    Screen.MousePointer = vbHourglass

    Con.ConnectionString = cCONNECT
    Con.CommandTimeout = 10000
    Con.Open
        Con.BeginTrans

        strSQL = "EXEC CN_CIERRE_DUAS '" & _
        Me.txtAno & "','" & _
        Me.txtMes & "'"
        
        ExecuteCommandSQL cCONNECT, strSQL

        Con.CommitTrans
        Screen.MousePointer = vbDefault
        MsgBox "Los datos fueron procesados con éxito.", vbInformation, "Mensaje"

    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    Screen.MousePointer = vbDefault
    ErrorHandler err, "Salvar_Datos"
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         Me.txtMes.SetFocus
         SelectionText Me.txtMes
    End If
End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.FunctButt2.SetFocus
    End If
End Sub



