VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_ShowVacaciones_Prog 
   Caption         =   "Autorizar De Vacaciones"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   8040
      TabIndex        =   6
      Top             =   6720
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"Frm_ShowVacaciones_Prog.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin VB.CommandButton Cmd_buscar 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   9120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Txt_Codigo 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Txt_Descripcion 
         Height          =   285
         Left            =   2760
         TabIndex        =   3
         Top             =   600
         Width           =   2295
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5655
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   9975
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "Frm_ShowVacaciones_Prog.frx":0093
         Column(2)       =   "Frm_ShowVacaciones_Prog.frx":015B
         FormatStylesCount=   6
         FormatStyle(1)  =   "Frm_ShowVacaciones_Prog.frx":01FF
         FormatStyle(2)  =   "Frm_ShowVacaciones_Prog.frx":0337
         FormatStyle(3)  =   "Frm_ShowVacaciones_Prog.frx":03E7
         FormatStyle(4)  =   "Frm_ShowVacaciones_Prog.frx":049B
         FormatStyle(5)  =   "Frm_ShowVacaciones_Prog.frx":0573
         FormatStyle(6)  =   "Frm_ShowVacaciones_Prog.frx":062B
         ImageCount      =   0
         PrinterProperties=   "Frm_ShowVacaciones_Prog.frx":070B
      End
      Begin MSComCtl2.DTPicker dtpAnoMes 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM yyyy"
         Format          =   23658499
         CurrentDate     =   37887
      End
      Begin VB.Label Label2 
         Caption         =   "Año-Mes"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro De Costo"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1170
      End
   End
End
Attribute VB_Name = "Frm_ShowVacaciones_Prog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Codigo As String
Public Descripcion As String

Private Sub Cmd_buscar_Click()
mostrar
End Sub

Private Sub Form_Load()
dtpAnoMes.Value = Date
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "AUTORIZAR"
        If GridEX1.RowCount = 0 Then Exit Sub
        Grabar
    Case "SALIR"
        Unload Me
    
End Select
End Sub

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)

If GridEX1.Columns("Fec_Prev_Vacaciones").Index = ColIndex Then
    If Format(GridEX1.Value(GridEX1.Columns("Fec_Prev_Vacaciones").Index), "MM yyyy") < Format(dtpAnoMes.Value, "MM yyyy") Then
        GridEX1.Value(GridEX1.Columns("Fec_Prev_Vacaciones").Index) = ""
    End If
End If

Select Case ColIndex
   Case Is = GridEX1.Columns("Fec_Prev_Vacaciones").Index
        Cancel = False
   Case Else
         Cancel = True
End Select
    

End Sub


Private Sub GridEX1_BeforeColUpdate(ByVal Row As Long, ByVal ColIndex As Integer, ByVal OldValue As String, ByVal Cancel As GridEX20.JSRetBoolean)
If GridEX1.Columns("Fec_Prev_Vacaciones").Index = ColIndex Then
    If Format(GridEX1.Value(GridEX1.Columns("Fec_Prev_Vacaciones").Index), "MM yyyy") < Format(dtpAnoMes.Value, "MM yyyy") Then
        GridEX1.Value(GridEX1.Columns("Fec_Prev_Vacaciones").Index) = ""
    End If
End If

End Sub

Private Sub Txt_Codigo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        If Trim(Me.Txt_Codigo.Text) = "" Then
            Call Me.BUSCA_CCOSTO(1)
        Else
            Call Me.BUSCA_CCOSTO(3)
        End If
    End If
End Sub


Sub BUSCA_CCOSTO(ByVal tipo As String)
Dim oTipo As New frmBusqGeneral
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    If tipo = "1" Then
        oTipo.sQuery = "Rh_muestra_CentroCosto_Ano_Mes '" & Format(dtpAnoMes, "yyyy") & "','" & Format(dtpAnoMes, "mm") & "','" & Txt_T & "'"
    Else
         oTipo.sQuery = "Rh_muestra_CentroCosto_Ano_Mes '" & Format(dtpAnoMes, "yyyy") & "','" & Format(dtpAnoMes, "mm") & "','" & Txt_Tipo & "'"
    End If
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        Me.Txt_Codigo = Trim(Codigo)
        Me.Txt_Descripcion = Trim(Descripcion)
        Codigo = ""
        Descripcion = ""
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub

Sub mostrar()
Dim strSQL As String
On Error GoTo Fin

    strSQL = "Rh_Sm_Muestra_Vacaciones '" & Format(dtpAnoMes, "yyyy") & "','" & Format(dtpAnoMes, "mm") & "','" & Txt_Codigo & "'"

    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
    
    GridEX1.Columns("Codigo").Width = 800
    GridEX1.Columns("Codigo").Caption = "Codigo"
    GridEX1.Columns("Codigo").HeaderAlignment = jgexAlignCenter
    
    
    GridEX1.Columns("fec_ingreso").Width = 1000
    GridEX1.Columns("fec_ingreso").Caption = "Fec Ingreso"
    GridEX1.Columns("fec_ingreso").HeaderAlignment = jgexAlignCenter
    
    
    GridEX1.Columns("Apellido_Paterno").Width = 1900
    GridEX1.Columns("Apellido_Paterno").Caption = "Apell Paterno"
    GridEX1.Columns("Apellido_Paterno").HeaderAlignment = jgexAlignCenter
    
    
    GridEX1.Columns("Apellido_Materno").Width = 1800
    GridEX1.Columns("Apellido_Materno").Caption = "Apell Materno"
    GridEX1.Columns("Apellido_Materno").HeaderAlignment = jgexAlignCenter

      
    GridEX1.Columns("Nombre_Trabajador").Width = 1500
    GridEX1.Columns("Nombre_Trabajador").Caption = "Nombre"
    GridEX1.Columns("Nombre_Trabajador").HeaderAlignment = jgexAlignCenter
    
   
    
     GridEX1.Columns("Fec_Prev_Vacaciones").Width = 1500
    GridEX1.Columns("Fec_Prev_Vacaciones").Caption = "Fec Prev Vacaciones"
    GridEX1.Columns("Fec_Prev_Vacaciones").HeaderAlignment = jgexAlignCenter

    
    
    GridEX1.Columns("Cod_CenCost").Width = 0
      
    
Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

Public Sub Grabar()
Dim sRows As Integer
On Error GoTo hand

    GridEX1.Row = 1
    
        For i = 1 To GridEX1.RowCount
        
        If GridEX1.Value(GridEX1.Columns("Fec_Prev_Vacaciones").Index) <> "" Then
            strSQL = "Rh_Up_Genera_Vacaciones '" & Format(dtpAnoMes, "yyyy") & "','" & Format(dtpAnoMes, "mm") & "','001','" & _
            GridEX1.Value(GridEX1.Columns("Tip_Trabajador").Index) & "','" & _
            GridEX1.Value(GridEX1.Columns("Cod_Trabajador").Index) & "','" & _
            GridEX1.Value(GridEX1.Columns("Fec_Prev_Vacaciones").Index) & "'"
        
            Call ExecuteSQL(cCONNECT, strSQL)
        End If

            GridEX1.MoveNext
        Next
    
Exit Sub
hand:
    ErrorHandler Err, "SALVAR"
    Set GridEX1.ADORecordset = Nothing
End Sub


