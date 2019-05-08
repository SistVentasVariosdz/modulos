VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmReclamosAviosProd 
   Caption         =   "Reclamos Avios Produccion Almacén"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir Boleta de Reclamo"
      Height          =   510
      Left            =   7665
      TabIndex        =   52
      Top             =   7725
      Width           =   1290
   End
   Begin VB.TextBox txtNum_Numerado 
      Height          =   315
      Left            =   1950
      TabIndex        =   51
      Top             =   4605
      Width           =   1080
   End
   Begin VB.OptionButton optOP 
      Caption         =   "N/P"
      Height          =   210
      Left            =   3195
      TabIndex        =   50
      Top             =   135
      Width           =   600
   End
   Begin VB.TextBox txtAbr_Fabrica2 
      Height          =   285
      Left            =   4005
      MaxLength       =   5
      TabIndex        =   48
      Top             =   480
      Width           =   630
   End
   Begin VB.TextBox txtNom_Fabrica2 
      Height          =   285
      Left            =   4950
      TabIndex        =   47
      Top             =   480
      Width           =   2850
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   330
      Left            =   4635
      TabIndex        =   46
      Top             =   465
      Width           =   330
   End
   Begin VB.TextBox txtOp2 
      Height          =   285
      Left            =   4005
      MaxLength       =   5
      TabIndex        =   45
      Top             =   105
      Width           =   735
   End
   Begin VB.TextBox txtDes_Estpro2 
      Height          =   300
      Left            =   4770
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   105
      Width           =   3030
   End
   Begin VB.TextBox txtCorrelativo 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   4635
      MaxLength       =   5
      TabIndex        =   1
      Top             =   3810
      Width           =   1530
   End
   Begin VB.OptionButton optFecha 
      Caption         =   "Fecha"
      Height          =   210
      Left            =   255
      TabIndex        =   41
      Top             =   195
      Value           =   -1  'True
      Width           =   1005
   End
   Begin VB.TextBox txtObservacion 
      Height          =   315
      Left            =   1095
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   7245
      Width           =   8475
   End
   Begin VB.Frame Fradetalle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   150
      TabIndex        =   30
      Tag             =   "Detail"
      Top             =   5340
      Width           =   10095
      Begin VB.TextBox TxtDesitem 
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   300
         Width           =   3345
      End
      Begin VB.TextBox TxtItem 
         BackColor       =   &H00FFFFFF&
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
         Left            =   930
         MaxLength       =   8
         TabIndex        =   9
         Top             =   300
         Width           =   945
      End
      Begin VB.TextBox CmbColor 
         BackColor       =   &H00FFFFFF&
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
         Left            =   945
         MaxLength       =   7
         TabIndex        =   13
         Top             =   630
         Width           =   945
      End
      Begin VB.TextBox TxtDetalle 
         Height          =   315
         Left            =   1920
         TabIndex        =   14
         Top             =   630
         Width           =   3315
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Left            =   930
         TabIndex        =   19
         Text            =   "0"
         Top             =   1290
         Width           =   945
      End
      Begin VB.TextBox TxtDes_Destino 
         Height          =   315
         Left            =   1920
         TabIndex        =   17
         Top             =   960
         Width           =   3315
      End
      Begin VB.TextBox Txtcod_Destino 
         BackColor       =   &H00FFFFFF&
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
         Left            =   930
         MaxLength       =   7
         TabIndex        =   16
         Top             =   960
         Width           =   945
      End
      Begin VB.TextBox TxtCod_Comb 
         BackColor       =   &H00FFFFFF&
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
         Left            =   6375
         MaxLength       =   8
         TabIndex        =   11
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox TxtDes_comb 
         Height          =   315
         Left            =   6960
         TabIndex        =   12
         Top             =   300
         Width           =   2985
      End
      Begin VB.TextBox TxtDes_Medida 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   7200
         TabIndex        =   31
         Top             =   630
         Width           =   2745
      End
      Begin VB.TextBox TxtCod_Talla 
         BackColor       =   &H00FFFFFF&
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
         Left            =   6360
         MaxLength       =   8
         TabIndex        =   15
         Top             =   630
         Width           =   825
      End
      Begin VB.TextBox TxtCod_EstCli 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Left            =   6360
         TabIndex        =   18
         Top             =   960
         Width           =   2625
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   38
         Top             =   1365
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Combinacion:"
         Height          =   195
         Index           =   2
         Left            =   5400
         TabIndex        =   37
         Top             =   405
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   36
         Top             =   405
         Width           =   345
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
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
         Index           =   3
         Left            =   240
         TabIndex        =   35
         Tag             =   "Hilado :"
         Top             =   1035
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estilo:"
         Height          =   195
         Index           =   3
         Left            =   5400
         TabIndex        =   34
         Top             =   1035
         Width           =   420
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Color"
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
         Index           =   5
         Left            =   240
         TabIndex        =   33
         Tag             =   "Hilado :"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Talla:"
         Height          =   195
         Index           =   4
         Left            =   5370
         TabIndex        =   32
         Top             =   720
         Width           =   390
      End
   End
   Begin VB.TextBox TxtCod_Motivo_Reclamo 
      Height          =   285
      Left            =   1005
      TabIndex        =   7
      Top             =   5010
      Width           =   900
   End
   Begin VB.TextBox TxtDes_Motivo_Reclamo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1950
      TabIndex        =   8
      Top             =   5010
      Width           =   3105
   End
   Begin VB.CommandButton CmdNuevo_MotRechazo 
      Caption         =   "..."
      Height          =   330
      Left            =   5205
      TabIndex        =   28
      Top             =   5010
      Width           =   540
   End
   Begin VB.TextBox txtNum_SecOrd 
      Height          =   315
      Left            =   1020
      TabIndex        =   6
      Top             =   4605
      Width           =   630
   End
   Begin VB.TextBox txtAbr_Fabrica 
      Height          =   285
      Left            =   1005
      MaxLength       =   5
      TabIndex        =   2
      Top             =   4230
      Width           =   630
   End
   Begin VB.TextBox txtNom_Fabrica 
      Height          =   285
      Left            =   1950
      TabIndex        =   3
      Top             =   4230
      Width           =   1800
   End
   Begin VB.CommandButton cmdBuscaFabrica 
      Caption         =   "..."
      Height          =   285
      Left            =   1635
      TabIndex        =   24
      Top             =   4230
      Width           =   300
   End
   Begin VB.TextBox TxtOP 
      Height          =   285
      Left            =   4305
      MaxLength       =   5
      TabIndex        =   4
      Top             =   4230
      Width           =   735
   End
   Begin VB.TextBox TxtEstilo 
      Height          =   300
      Left            =   5070
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4230
      Width           =   3030
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   300
      Left            =   1335
      TabIndex        =   23
      Top             =   90
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   529
      _Version        =   393216
      Format          =   23920641
      CurrentDate     =   39001
   End
   Begin GridEX20.GridEX gex 
      Height          =   2715
      Left            =   135
      TabIndex        =   40
      Top             =   930
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   4789
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorBkg    =   -2147483624
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmReclamosAviosProd.frx":0000
      FormatStyle(2)  =   "frmReclamosAviosProd.frx":0138
      FormatStyle(3)  =   "frmReclamosAviosProd.frx":01E8
      FormatStyle(4)  =   "frmReclamosAviosProd.frx":029C
      FormatStyle(5)  =   "frmReclamosAviosProd.frx":0374
      FormatStyle(6)  =   "frmReclamosAviosProd.frx":042C
      FormatStyle(7)  =   "frmReclamosAviosProd.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frmReclamosAviosProd.frx":052C
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   600
      Left            =   3405
      TabIndex        =   21
      Top             =   7650
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmReclamosAviosProd.frx":0704
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin MSComCtl2.DTPicker dtpFec_Solicitud 
      Height          =   345
      Left            =   1005
      TabIndex        =   0
      Top             =   3810
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   609
      _Version        =   393216
      Format          =   23920641
      CurrentDate     =   39001
   End
   Begin FunctionsButtons.FunctButt funcBuscar 
      Height          =   480
      Left            =   8970
      TabIndex        =   43
      Top             =   15
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   847
      Custom          =   "0~0~BUSCAR~True~True~&Buscar~0~0~4~~0~True~False~&Buscar~"
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1000
      ControlHeigth   =   450
      ControlSeparator=   80
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   7725
      Top             =   4935
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Fabrica"
      Height          =   195
      Left            =   3450
      TabIndex        =   49
      Top             =   510
      Width           =   525
   End
   Begin VB.Label Label1 
      Caption         =   "Solicitud"
      Height          =   225
      Index           =   5
      Left            =   3930
      TabIndex        =   42
      Top             =   3855
      Width           =   630
   End
   Begin VB.Label Label13 
      Caption         =   "Observacion:"
      Height          =   450
      Left            =   165
      TabIndex        =   39
      Top             =   7320
      Width           =   960
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Motivo"
      Height          =   195
      Left            =   195
      TabIndex        =   29
      Top             =   5040
      Width           =   480
   End
   Begin VB.Label lblsecuencia 
      AutoSize        =   -1  'True
      Caption         =   "O/Corte :"
      Height          =   195
      Left            =   195
      TabIndex        =   27
      Top             =   4665
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fabrica"
      Height          =   195
      Left            =   180
      TabIndex        =   26
      Top             =   4260
      Width           =   525
   End
   Begin VB.Label Label2 
      Caption         =   "OP"
      Height          =   240
      Left            =   3960
      TabIndex        =   25
      Top             =   4275
      Width           =   345
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   225
      Index           =   0
      Left            =   195
      TabIndex        =   22
      Top             =   3840
      Width           =   540
   End
End
Attribute VB_Name = "frmReclamosAviosProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL  As String
Dim vMessage As String
Public Codigo As String
Public Descripcion As String
Public TipoAdd As String
Dim sAccion As String
Dim Nivel As String
Public varCod_Fabrica As String
Public varCod_Fabrica2 As String
Public varNum_SecOrd As String
Public PASO As Boolean

Private Sub cmdImprimir_Click()
    Reporte
End Sub

Private Sub dtpFec_Solicitud_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    VerificaFabrica txtAbr_Fabrica, txtNom_Fabrica, varCod_Fabrica
    VerificaFabrica txtAbr_Fabrica2, txtNom_Fabrica2, varCod_Fabrica2
    DESHABILITA
    MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp1, Me.Name)
End Sub

Private Sub funcBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    carga_grid
End Sub

Private Sub optOP_Click()
    txtOp2.SetFocus
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
    Me.TxtOP.SetFocus
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtObservacion.SetFocus
    End If
End Sub

Private Sub TxtItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BUSCAENVIADO_PRODUCCION
    End If
End Sub

Private Sub txtNom_Fabrica_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Me.BUSCA_FABRICA(2)
    End If
End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        MantFunc1.SetFocus
    End If
End Sub

Private Sub TxtOP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Dim sCod_Fabrica As String

        strSQL = "select cod_fabrica from tg_fabrica where abr_fabrica='" & Me.txtAbr_Fabrica.Text & "'"
        varCod_Fabrica = DevuelveCampo(strSQL, cConnect)

        TxtOP.Text = Format(Trim(TxtOP.Text), "00000")
        If DevuelveCampo("select count(*) from es_Ordpro where cod_fabrica='" & varCod_Fabrica & "' AND cod_ordpro = '" & TxtOP.Text & "'", cConnect) > 0 Then
            strSQL = "SELECT cod_fabrica FROM TG_FABRICA WHERE Abr_Fabrica = '" & Trim(Me.txtAbr_Fabrica.Text) & "'"
            Me.TxtEstilo.Text = DevuelveCampo("SELECT b.Des_EstPro FROM   ES_OrdPro  a , ES_EstPRo b WHERE  a.Cod_EstPro = b.Cod_EstPRo AND a.Cod_Fabrica= '" & DevuelveCampo(strSQL, cConnect) & "' AND a.Cod_OrdPro = '" & TxtOP.Text & "'", cConnect)
            txtNum_SecOrd.SetFocus
        Else
            MsgBox "Codigo de N/P no existe", vbInformation, Me.Caption
        End If
    End If
End Sub


Private Sub txtNum_SecOrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.txtNum_SecOrd.Text = "" Then
            Call BUSCA_Num_SecOrd
            TxtCod_Motivo_Reclamo.SetFocus
        End If
        
    End If
End Sub


Public Sub BUSCA_Num_SecOrd()
    
    Dim oTipo As New frmBusqNum_SecOrd
    Dim rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    
    oTipo.sQuery = "EXEC UP_SEL_CFORDPRO_Num_SecOrd '" & varCod_Fabrica & "','" & Me.TxtOP.Text & "'"
    
    oTipo.CARGAR_DATOS
    oTipo.Show 1
    If varNum_SecOrd <> "" Then
        Me.txtNum_SecOrd.Text = varNum_SecOrd
        varNum_SecOrd = ""
    End If
    Set oTipo = Nothing
    Set rs = Nothing
    
    
    
End Sub



Private Sub TxtCod_Motivo_Reclamo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call AYUDA_Motivo_Reclamo
    If TxtItem.Enabled Then
        TxtItem.SetFocus
    End If
End If
End Sub

Sub AYUDA_Motivo_Reclamo()
    Dim oTipo As New frmBusqGeneral2
    Dim rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    Codigo = ""
    Descripcion = ""
    oTipo.sQuery = "SELECT Cod_Motivo_Reclamo AS 'Código', Des_Motivo_Reclamo as 'Descripción' FROM LG_TIPOS_RECLAMOS_AVIOS_PRODUCCION_ALMACEN order by 1"
    oTipo.CARGAR_DATOS
    oTipo.Show 1
    If Codigo <> "" Then
        Me.TxtCod_Motivo_Reclamo.Text = Trim(Codigo)
        Me.TxtDes_Motivo_Reclamo.Text = Trim(Descripcion)
        Codigo = "": Descripcion = ""
        'txtCod_TemCli.SetFocus
    End If
    Set oTipo = Nothing
    Set rs = Nothing

End Sub


Sub carga_grid()
On Error GoTo ErrCargaGrid
    Dim vBookmark As Variant
    Dim sopcion As String
    
    vBookmark = gex.Row
    
    If OptFecha.Value Then
        sopcion = "1"
    Else
        sopcion = "2"
    End If
    
    strSQL = "lg_muestra_reclamos_avios '$','$' ,'$','$'"
    strSQL = VBsprintf(strSQL, sopcion, DTPFecha.Value, varCod_Fabrica2, txtOp2.Text)
    
    Set gex.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
    gex.Row = vBookmark
    
    
    Exit Sub
ErrCargaGrid:
ErrorHandler err, "Carga_Grid"
End Sub




Private Sub gex_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
If gex.RowCount > 0 Then
    If gex.Value(gex.Columns("Fec_Solicitud").Index) <> Empty Then
        dtpFec_Solicitud.Value = gex.Value(gex.Columns("Fec_Solicitud").Index)
    End If
    txtAbr_Fabrica.Text = gex.Value(gex.Columns("Abr_Fabrica").Index)
    txtNom_Fabrica.Text = gex.Value(gex.Columns("Nom_Fabrica").Index)
    txtCorrelativo.Text = gex.Value(gex.Columns("Correlativo").Index)
    TxtOP.Text = gex.Value(gex.Columns("Cod_OrdPro").Index)
    TxtEstilo.Text = gex.Value(gex.Columns("des_estpro").Index)
    txtNum_SecOrd.Text = gex.Value(gex.Columns("num_secord").Index)
    TxtCod_Motivo_Reclamo.Text = gex.Value(gex.Columns("Cod_Motivo_Reclamo").Index)
    TxtDes_Motivo_Reclamo.Text = gex.Value(gex.Columns("Des_Motivo_Reclamo").Index)
    TxtItem.Text = gex.Value(gex.Columns("cod_item").Index)
    TxtDesitem.Text = gex.Value(gex.Columns("des_item").Index)
    CmbColor.Text = gex.Value(gex.Columns("cod_color").Index)
    TxtDetalle.Text = gex.Value(gex.Columns("des_color").Index)
    Txtcod_Destino.Text = gex.Value(gex.Columns("cod_destino").Index)
    TxtDes_Destino.Text = gex.Value(gex.Columns("des_destino").Index)
    TxtCantidad.Text = gex.Value(gex.Columns("can_requerida").Index)
    
    TxtCod_Comb.Text = gex.Value(gex.Columns("cod_comb").Index)
    TxtDes_comb.Text = gex.Value(gex.Columns("des_comb").Index)
    TxtCod_Talla.Text = gex.Value(gex.Columns("cod_talla").Index)
    
    TxtCod_EstCli.Text = gex.Value(gex.Columns("cod_estcli").Index)
    
    txtObservacion.Text = gex.Value(gex.Columns("Observacion").Index)
    txtNum_Numerado.Text = gex.Value(gex.Columns("Num_Numerado").Index)
    varCod_Fabrica = gex.Value(gex.Columns("cod_fabrica").Index)
End If
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ADICIONAR"
            sAccion = "I"
            LIMPIA_DATOS
            
            HABILITA
                
             dtpFec_Solicitud.Enabled = True
             TxtOP.Enabled = True
             TxtEstilo.Enabled = True
             txtNum_SecOrd.Enabled = True
             TxtCod_Motivo_Reclamo.Enabled = True
             TxtDes_Motivo_Reclamo.Enabled = True
             TxtItem.Enabled = True
             TxtDesitem.Enabled = True
             CmbColor.Enabled = True
             TxtDetalle.Enabled = True
             Txtcod_Destino.Enabled = True
             TxtDes_Destino.Enabled = True
            
             TxtCod_Comb.Enabled = True
             TxtDes_comb.Enabled = True
             TxtCod_Talla.Enabled = True
             TxtDes_Medida.Enabled = True
             TxtCod_EstCli.Enabled = True
             
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            dtpFec_Solicitud.SetFocus
        Case "MODIFICAR"
            If gex.RowCount = 0 Then Exit Sub
            sAccion = "U"
            HABILITA
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Case "ELIMINAR"
            
            If gex.RowCount = 0 Then Exit Sub
            vMessage = MsgBox("Esta seguro que desea eliminar el registro", vbYesNo, "Eliminar")
            If vMessage = vbYes Then
                sAccion = "D"
                SALVAR_DATOS
            End If
            carga_grid
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Case "GRABAR"
                
                 If SALVAR_DATOS = False Then Exit Sub
                 HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                 Me.MantFunc1.SetFocus
                 sAccion = ""
                 carga_grid
        Case "DESHACER"
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            carga_grid
            DESHABILITA
        Case "SALIR"
            Unload Me
    End Select
End Sub

Function SALVAR_DATOS() As Boolean
On Error GoTo ErrSalvarDatos
strSQL = "LG_MAN_Reclamos_Avios_Produccion_Almacen '$','$','$','$','$','$','$','$','$','$','$','$','$','$','$','$'"
strSQL = VBsprintf(strSQL, sAccion, txtCorrelativo.Text, dtpFec_Solicitud.Value, vusu, TxtCod_Motivo_Reclamo, varCod_Fabrica, TxtOP, txtNum_SecOrd.Text, TxtItem.Text, TxtCod_Comb, CmbColor, TxtCod_Talla, Txtcod_Destino.Text, TxtCod_EstCli.Text, txtObservacion.Text, TxtCantidad.Text)


ExecuteSQL cConnect, strSQL
SALVAR_DATOS = True
Exit Function

ErrSalvarDatos:
    SALVAR_DATOS = False
    ErrorHandler err, "SALVAR_DATOS"
End Function

Sub LIMPIA_DATOS()
    dtpFec_Solicitud.Value = Date
    txtCorrelativo.Text = ""
    TxtOP.Text = ""
    TxtEstilo.Text = ""
    txtNum_SecOrd.Text = ""
    TxtCod_Motivo_Reclamo.Text = ""
    TxtDes_Motivo_Reclamo.Text = ""
    TxtItem.Text = ""
    TxtDesitem.Text = ""
    CmbColor.Text = ""
    TxtDetalle.Text = ""
    Txtcod_Destino.Text = ""
    TxtDes_Destino.Text = ""
    TxtCantidad.Text = 0
    
    TxtCod_Comb.Text = ""
    TxtDes_comb.Text = ""
    TxtCod_Talla.Text = ""
    TxtDes_Medida.Text = ""
    TxtCod_EstCli.Text = ""
    txtNum_Numerado.Text = ""
    
    txtObservacion.Text = ""
      
End Sub


Sub DESHABILITA()
    dtpFec_Solicitud.Enabled = False
    txtAbr_Fabrica.Enabled = False
    txtNom_Fabrica.Enabled = False
    txtCorrelativo.Enabled = False
    TxtOP.Enabled = False
    TxtEstilo.Enabled = False
    txtNum_SecOrd.Enabled = False
    TxtCod_Motivo_Reclamo.Enabled = False
    TxtDes_Motivo_Reclamo.Enabled = False
    TxtItem.Enabled = False
    TxtDesitem.Enabled = False
    CmbColor.Enabled = False
    TxtDetalle.Enabled = False
    Txtcod_Destino.Enabled = False
    TxtDes_Destino.Enabled = False
    TxtCantidad.Enabled = False
    
    TxtCod_Comb.Enabled = False
    TxtDes_comb.Enabled = False
    TxtCod_Talla.Enabled = False
    TxtDes_Medida.Enabled = False
    TxtCod_EstCli.Enabled = False
    txtNum_Numerado.Enabled = False
    
    txtObservacion.Enabled = False

End Sub

Sub HABILITA()
    TxtCod_Motivo_Reclamo.Enabled = True
    TxtDes_Motivo_Reclamo.Enabled = True
    TxtCantidad.Enabled = True
    txtObservacion.Enabled = True
End Sub


Private Sub txtOp2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Dim sCod_Fabrica As String

        strSQL = "select cod_fabrica from tg_fabrica where abr_fabrica='" & Me.txtAbr_Fabrica2.Text & "'"
        varCod_Fabrica2 = DevuelveCampo(strSQL, cConnect)

        txtOp2.Text = Format(Trim(txtOp2.Text), "00000")
        If DevuelveCampo("select count(*) from es_Ordpro where cod_fabrica='" & varCod_Fabrica2 & "' AND cod_ordpro = '" & txtOp2.Text & "'", cConnect) > 0 Then
            strSQL = "SELECT cod_fabrica FROM TG_FABRICA WHERE Abr_Fabrica = '" & Trim(Me.txtAbr_Fabrica2.Text) & "'"
            Me.txtDes_Estpro2.Text = DevuelveCampo("SELECT b.Des_EstPro FROM   ES_OrdPro  a , ES_EstPRo b WHERE  a.Cod_EstPro = b.Cod_EstPRo AND a.Cod_Fabrica= '" & DevuelveCampo(strSQL, cConnect) & "' AND a.Cod_OrdPro = '" & txtOp2.Text & "'", cConnect)
            
        Else
            MsgBox "Codigo de N/P no existe", vbInformation, Me.Caption
        End If
    End If

End Sub


Private Sub BUSCAENVIADO_PRODUCCION()
On Error GoTo Fin
Dim rstAux As ADODB.Recordset
    
    strSQL = "ES_MUESTRA_ITEMS_ENVIADOS_PRODUCCION '$', '$'"
    strSQL = VBsprintf(strSQL, varCod_Fabrica, TxtOP.Text)
    
    With frmBusqGeneral4
        Set .oParent = Me
        .sQuery = strSQL
        .CARGAR_DATOS
        
        .DGridLista.Columns("COD_COMB").Width = 400
        .DGridLista.Columns("des_item").Width = 2000
        .DGridLista.Columns("DES_COMB").Width = 1000
        .DGridLista.Columns("COD_COLOR").Width = 800
        .DGridLista.Columns("COD_TALLA").Width = 800
        .DGridLista.Columns("COD_DESTINO").Width = 800
        .DGridLista.Columns("DES_DESTINO").Width = 800
        .DGridLista.Columns("COD_ESTCLI").Width = 800
        
        Codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            TxtItem = Trim(rstAux!cod_item)
            TxtDesitem = Trim(rstAux!des_item)
            TxtCod_Comb = Trim(rstAux!Cod_Comb)
            TxtDes_comb = Trim(rstAux!Des_Comb)
            CmbColor = Trim(rstAux!Cod_color)
            TxtDetalle = Trim(rstAux!Des_color)
            TxtCod_Talla = Trim(rstAux!Cod_Talla)
            Txtcod_Destino = Trim(rstAux!Cod_Destino)
            TxtDes_Destino = Trim(rstAux!Des_Destino)
            TxtCod_EstCli = Trim(rstAux!cod_estcli)
            TxtCantidad = Trim(rstAux!Cantidad)
            
            TxtCantidad.SetFocus
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    
Exit Sub
Resume
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda "
End Sub
    
Private Sub VerificaFabrica(ByRef objFabrica As TextBox, ByRef objNombreFabrica As TextBox, ByRef variable As String)
    Dim sSQL As String
    Dim iRet As String
    Dim rs As ADODB.Recordset
    
    sSQL = "SELECT count(*) FROM TG_Fabrica "
    iRet = DevuelveCampo(sSQL, cConnect)
    If iRet = 1 Then
        Set rs = GetRecordset(cConnect, "SELECT * FROM TG_Fabrica ")
        If Not rs Is Nothing Then
            objFabrica.Text = rs!Abr_Fabrica
            
            
            objNombreFabrica.Text = rs!Nom_Fabrica
            objFabrica.Enabled = False
            objNombreFabrica.Enabled = False
            
            variable = DevuelveCampo(sSQL, cConnect)
        End If
    End If
End Sub



Sub Reporte()
On Error GoTo err:
Dim oo As Object

If gex.RowCount = 0 Then Exit Sub

Set oo = CreateObject("excel.application")
oo.workbooks.Open vRuta & "\RptBoletaReclamoAvios.xlt"
oo.Visible = True
oo.run "Reporte", gex.Value(gex.Columns("correlativo").Index), cConnect
Set oo = Nothing
Exit Sub
err:
    MsgBox "Error en la Impresion del Boleta de Reclamo Avios"
End Sub

