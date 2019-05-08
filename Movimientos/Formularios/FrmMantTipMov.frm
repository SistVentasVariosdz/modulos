VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmMantTipMov 
   Caption         =   "Tipos de Movimiento"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
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
      Height          =   3150
      Left            =   90
      TabIndex        =   6
      Tag             =   "Detail"
      Top             =   3390
      Width           =   7800
      Begin VB.Frame Frame4 
         Caption         =   "2da Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3720
         TabIndex        =   36
         Top             =   2370
         Width           =   1995
         Begin VB.OptionButton optNo2da 
            Caption         =   "NO"
            Height          =   210
            Left            =   990
            TabIndex        =   38
            Top             =   300
            Value           =   -1  'True
            Width           =   570
         End
         Begin VB.OptionButton optSi2da 
            Caption         =   "SI"
            Height          =   210
            Left            =   135
            TabIndex        =   37
            Top             =   300
            Width           =   525
         End
      End
      Begin VB.ComboBox CmbFabrica 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   2490
         Width           =   2355
      End
      Begin VB.ComboBox CmbTipProd 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2130
         Width           =   2355
      End
      Begin VB.ComboBox CmbCalidad 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1740
         Width           =   2355
      End
      Begin VB.ComboBox CmbAnexo 
         Height          =   315
         Left            =   4755
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1410
         Width           =   2355
      End
      Begin VB.ComboBox CmbTipItem 
         Height          =   315
         Left            =   4740
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   660
         Width           =   2355
      End
      Begin VB.TextBox TxtCentCosto 
         Height          =   315
         Left            =   4740
         MaxLength       =   50
         TabIndex        =   26
         Top             =   1020
         Width           =   2340
      End
      Begin VB.TextBox TxtOP 
         Height          =   315
         Left            =   975
         MaxLength       =   50
         TabIndex        =   24
         Top             =   1380
         Width           =   2340
      End
      Begin VB.Frame Frame3 
         Caption         =   "Accion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3720
         TabIndex        =   21
         Top             =   1740
         Width           =   1995
         Begin VB.OptionButton OptAccionSi 
            Caption         =   "Interna"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton OptAccionNo 
            Caption         =   "Externa"
            Height          =   255
            Left            =   960
            TabIndex        =   22
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Valorizable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5760
         TabIndex        =   11
         Top             =   1740
         Width           =   1965
         Begin VB.OptionButton optNo 
            Caption         =   "NO"
            Height          =   255
            Left            =   1200
            TabIndex        =   13
            Top             =   240
            Width           =   645
         End
         Begin VB.OptionButton optSi 
            Caption         =   "SI"
            Height          =   255
            Left            =   270
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   465
         End
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   315
         Left            =   4740
         MaxLength       =   50
         TabIndex        =   10
         Top             =   300
         Width           =   2340
      End
      Begin VB.TextBox TxtCodigo 
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
         Left            =   1005
         MaxLength       =   3
         TabIndex        =   9
         Top             =   300
         Width           =   615
      End
      Begin VB.ComboBox CmbClaseMov 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   660
         Width           =   2355
      End
      Begin VB.ComboBox CmbClaseOC 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1020
         Width           =   2355
      End
      Begin VB.Label Label4 
         Caption         =   "Fabrica:"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   2580
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Prod:"
         Height          =   255
         Left            =   210
         TabIndex        =   33
         Top             =   2190
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "Calidad:"
         Height          =   225
         Left            =   210
         TabIndex        =   32
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anexo:"
         Height          =   195
         Index           =   5
         Left            =   3810
         TabIndex        =   29
         Top             =   1470
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo OP:"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   25
         Top             =   1425
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro Costo:"
         Height          =   195
         Index           =   3
         Left            =   3750
         TabIndex        =   20
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clase O/C"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   18
         Top             =   1065
         Width           =   735
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Clase Mov:"
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
         Left            =   195
         TabIndex        =   17
         Tag             =   "Hilado :"
         Top             =   735
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Item:"
         Height          =   195
         Index           =   1
         Left            =   3735
         TabIndex        =   16
         Top             =   765
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion:"
         Height          =   195
         Index           =   0
         Left            =   3735
         TabIndex        =   15
         Top             =   405
         Width           =   885
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
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
         Left            =   195
         TabIndex        =   14
         Tag             =   "Hilado :"
         Top             =   375
         Width           =   540
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
      Height          =   3255
      Left            =   60
      TabIndex        =   5
      Tag             =   "List"
      Top             =   60
      Width           =   7815
      Begin GridEX20.GridEX gexList 
         Height          =   2985
         Left            =   90
         TabIndex        =   39
         Top             =   180
         Width           =   7650
         _ExtentX        =   13494
         _ExtentY        =   5265
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         DataMode        =   1
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "FrmMantTipMov.frx":0000
         Column(2)       =   "FrmMantTipMov.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "FrmMantTipMov.frx":016C
         FormatStyle(2)  =   "FrmMantTipMov.frx":02A4
         FormatStyle(3)  =   "FrmMantTipMov.frx":0354
         FormatStyle(4)  =   "FrmMantTipMov.frx":0408
         FormatStyle(5)  =   "FrmMantTipMov.frx":04E0
         FormatStyle(6)  =   "FrmMantTipMov.frx":0598
         ImageCount      =   0
         PrinterProperties=   "FrmMantTipMov.frx":0678
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1245
      TabIndex        =   0
      Top             =   6600
      Width           =   1965
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "FrmMantTipMov.frx":0850
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Anterior"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "FrmMantTipMov.frx":09C2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Siguiente"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "FrmMantTipMov.frx":0B34
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Primero"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "FrmMantTipMov.frx":0CA6
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ultimo"
         Top             =   0
         Width           =   495
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   3180
      TabIndex        =   19
      Top             =   6570
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmMantTipMov.frx":0E18
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   300
      Top             =   6600
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmMantTipMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Estado As String
Dim StrSql As String
Sub Datos(Accion As String, EsAccion As Boolean)
Dim Reg As ADODB.Recordset
Set Reg = Nothing

Set Reg = New ADODB.Recordset
Reg.CursorLocation = adUseClient

Reg.Open "UP_Lg_TiposMov '" & Accion & "','" & TxtCodigo.Text & "','" & Right(Me.CmbClaseMov, 1) & "','" & TxtDescripcion & "','" & _
        Trim(Right(CmbTipItem, 2)) & "','" & Trim(Right(Me.CmbClaseOC, 2)) & "','" & Me.TxtCentCosto & "','" & _
        IIf(optSi, "S", "N") & "','" & IIf(OptAccionSi, "I", "E") & "','" & Me.TxtOP & "','" & Trim(Right(Me.CmbAnexo, 1)) & "','" & _
        Trim(Right(Me.CmbCalidad, 1)) & "','" & Trim(Right(Me.CmbTipProd, 2)) & "','" & Trim(Right(Me.CmbFabrica, 3)) & "','" & IIf(optSi2da, "S", "N") & "'", cConnect
            
CARGA_GRID

End Sub

Sub CARGA_GRID()
StrSql = "exec UP_Lg_TiposMov 'V','" & TxtCodigo.Text & "','" & Right(Me.CmbClaseMov, 1) & "','" & TxtDescripcion & "','" & _
        Trim(Right(CmbTipItem, 2)) & "','" & Trim(Right(Me.CmbClaseOC, 2)) & "','" & Me.TxtCentCosto & "','" & _
        IIf(optSi, "S", "N") & "','" & IIf(OptAccionSi, "I", "E") & "','" & Me.TxtOP & "','" & Trim(Right(Me.CmbAnexo, 1)) & "','" & _
        Trim(Right(Me.CmbCalidad, 1)) & "','" & Trim(Right(Me.CmbTipProd, 2)) & "','" & Trim(Right(Me.CmbFabrica, 3)) & "','" & IIf(optSi2da, "S", "N") & "'"
            

Set gexList.ADORecordset = CargarRecordSetDesconectado(StrSql, cConnect)

End Sub

Sub Habilita()
Me.TxtCentCosto.Enabled = True
Me.TxtCodigo.Enabled = True
Me.TxtDescripcion.Enabled = True
Me.TxtOP.Enabled = True
Me.CmbClaseMov.Enabled = True
Me.CmbClaseOC.Enabled = True
Me.CmbTipItem.Enabled = True
Me.CmbTipProd.Enabled = True
Me.CmbFabrica.Enabled = True
Me.CmbCalidad.Enabled = True
Me.CmbAnexo.Enabled = True

Frame2.Enabled = True
Frame3.Enabled = True
Frame4.Enabled = True
End Sub
Sub Deshabilita()
Me.TxtCentCosto.Enabled = False
Me.TxtCodigo.Enabled = False
Me.TxtDescripcion.Enabled = False
Me.TxtOP.Enabled = False
Me.CmbClaseMov.Enabled = False
Me.CmbClaseOC.Enabled = False
Me.CmbTipItem.Enabled = False
Me.CmbTipProd.Enabled = False
Me.CmbFabrica.Enabled = False
Me.CmbCalidad.Enabled = False
Me.CmbAnexo.Enabled = False

Frame2.Enabled = False
Frame3.Enabled = False
Frame4.Enabled = False
End Sub

Sub Limpia()
Me.TxtCentCosto = ""
Me.TxtCodigo = ""
Me.TxtDescripcion = ""
Me.TxtOP = ""
Me.CmbClaseMov.ListIndex = -1
Me.CmbClaseOC.ListIndex = -1
Me.CmbTipItem.ListIndex = -1
Me.CmbCalidad.ListIndex = -1
Me.CmbTipProd.ListIndex = -1
Me.CmbFabrica.ListIndex = -1

Me.optNo2da.Value = True
End Sub


Private Sub cmdFirst_Click()
    If gexList.RowCount > 0 Then gexList.MoveFirst
End Sub

Private Sub cmdLast_Click()
    If gexList.RowCount > 0 Then gexList.MoveLast
End Sub

Private Sub cmdNext_Click()
If gexList.RowCount > 0 Then gexList.MoveNext
End Sub

Private Sub cmdPrevious_Click()
If gexList.RowCount > 0 Then gexList.MoveLast
End Sub


Private Sub Form_Load()
'cCONNECT = "Provider=sqloledb;Server=SERVIDOR;Database=lives;UID=sa;pwd=;"
'cSEGURIDAD = "Provider=sqloledb;Server=servidor;Database=seguridad;UID=sa;pwd=;"

Limpia
Deshabilita
LlenaCombo CmbTipItem, "select Des_TipItem +space(100)+Tip_Item from lg_tipitem order by 1", cConnect
LlenaCombo CmbClaseMov, "select Des_ClaMov  +space(100)+Cod_ClaMov  from lg_clamov order by 1", cConnect
LlenaCombo Me.CmbClaseOC, "select Des_ClaOrdComp  +space(100)+Cod_ClaOrdComp  from lg_claordcomp order by 1", cConnect

LlenaCombo Me.CmbAnexo, "select DES_TIPANEX +space(100)+COD_TIPANEX from CN_TipoAnexoContable order by 1", cConnect
CmbAnexo.AddItem "[Ninguno]                                 ", 0

LlenaCombo Me.CmbCalidad, "select Descripcion +space(100)+Cod_calidad from tg_calidad order by 1", cConnect
LlenaCombo Me.CmbTipProd, "select Des_MpTp +space(100)+Tip_PtMp from tg_tipospr order by 1", cConnect
LlenaCombo Me.CmbFabrica, "select Des_Fabrica +space(100)+Cod_Fabrica from tg_fabrica order by 1", cConnect

'FormateaGrid Me.DGridLista
MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp1, Me.Name)
CARGA_GRID
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub gexList_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
If gexList.RowCount > 0 Then
        TxtCodigo = gexList.Value(gexList.Columns("Cod. Tipo").Index)
        TxtDescripcion = gexList.Value(gexList.Columns("Descripcion").Index)
        Me.TxtOP = gexList.Value(gexList.Columns("Tipo OP").Index)
        Me.TxtCentCosto = gexList.Value(gexList.Columns("Cent. Costo").Index)
        
        BuscaCombo gexList.Value(gexList.Columns("Clase Mov").Index), 1, Me.CmbClaseMov
        BuscaCombo gexList.Value(gexList.Columns("Item").Index), 1, Me.CmbTipItem
        BuscaCombo gexList.Value(gexList.Columns("Clase OC").Index), 1, Me.CmbClaseOC
        BuscaCombo gexList.Value(gexList.Columns("anexo").Index), 1, Me.CmbAnexo
        BuscaCombo gexList.Value(gexList.Columns("Calidad").Index), 1, Me.CmbCalidad
        BuscaCombo gexList.Value(gexList.Columns("Tipo Prod.").Index), 1, Me.CmbTipProd
        BuscaCombo gexList.Value(gexList.Columns("Fabrica").Index), 1, Me.CmbFabrica
        
    
        If gexList.Value(gexList.Columns("Valorizable").Index) = "SI" Then
            optSi.Value = True
        Else
            optNo.Value = True
        End If
        
        If gexList.Value(gexList.Columns("Accion").Index) = "I" Then
            OptAccionSi.Value = True
        Else
            OptAccionNo.Value = True
        End If
        
        If gexList.Value(gexList.Columns("2da.Numeracion").Index) = "S" Then
            optSi2da.Value = True
        Else
            optNo2da.Value = True
        End If
        
End If
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand
Select Case ActionName
    Case "ADICIONAR"
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Limpia
        Habilita
        Estado = "NUEVO"
        TxtCodigo.SetFocus
    Case "MODIFICAR"
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Estado = "MODIFICAR"
        Habilita
        TxtCodigo.Enabled = False
        TxtDescripcion.SetFocus
    Case "ELIMINAR"
        Datos "b", True
        Limpia
        CARGA_GRID
        Deshabilita
    Case "GRABAR"
        If CmbClaseMov = "" Then MsgBox "Seleccione una clase de movimiento", vbInformation: Exit Sub
        If CmbTipItem = "" Then MsgBox "Seleccione una clase de item", vbInformation: Exit Sub
        
        If Estado = "NUEVO" Then
            Datos "i", True
        Else
            Datos "a", True
        End If
        Limpia
        Deshabilita
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        CARGA_GRID
    Case "DESHACER"
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Limpia
        CARGA_GRID
        Deshabilita
    Case "SALIR"
        Unload Me
End Select

Exit Sub
hand:
ErrorHandler Err, "MantFunc1_ActionClick"
End Sub


Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtCodigo = Trim(DevuelveCampo("Select dbo.uf_devuelvecodigo(6," & IIf(Trim(TxtCodigo) = "", 0, TxtCodigo) & ")", cConnect))
End If
End Sub


