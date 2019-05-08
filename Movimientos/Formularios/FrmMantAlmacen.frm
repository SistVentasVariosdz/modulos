VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmMantAlmacen 
   Caption         =   "Mantenimiento de Almacen"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   510
      TabIndex        =   15
      Top             =   5385
      Width           =   1965
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "FrmMantAlmacen.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Ultimo"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "FrmMantAlmacen.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Primero"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "FrmMantAlmacen.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Siguiente"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "FrmMantAlmacen.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Anterior"
         Top             =   0
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
      Height          =   3255
      Left            =   90
      TabIndex        =   13
      Tag             =   "List"
      Top             =   45
      Width           =   7815
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2775
         Left            =   180
         TabIndex        =   14
         Top             =   345
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   4895
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
      Height          =   1890
      Left            =   75
      TabIndex        =   0
      Tag             =   "Detail"
      Top             =   3315
      Width           =   7800
      Begin VB.ComboBox cmbPresentacion 
         Height          =   315
         Left            =   4740
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1020
         Width           =   2355
      End
      Begin VB.ComboBox CmbTipItem 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   660
         Width           =   2355
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
         MaxLength       =   2
         TabIndex        =   6
         Top             =   300
         Width           =   615
      End
      Begin VB.TextBox TxtAlmacen 
         Height          =   315
         Left            =   4740
         MaxLength       =   50
         TabIndex        =   5
         Top             =   300
         Width           =   2340
      End
      Begin VB.TextBox TxtCuenta 
         Height          =   315
         Left            =   4740
         MaxLength       =   14
         TabIndex        =   4
         Top             =   660
         Width           =   2340
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operativo"
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
         Left            =   960
         TabIndex        =   1
         Top             =   1080
         Width           =   2145
         Begin VB.OptionButton optSi 
            Caption         =   "SI"
            Height          =   255
            Left            =   270
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   465
         End
         Begin VB.OptionButton optNo 
            Caption         =   "NO"
            Height          =   255
            Left            =   1200
            TabIndex        =   2
            Top             =   240
            Width           =   645
         End
      End
      Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
         Left            =   7290
         Top             =   1350
         _cx             =   847
         _cy             =   847
         PassiveMode     =   0   'False
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
         TabIndex        =   12
         Tag             =   "Hilado :"
         Top             =   375
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Almacen"
         Height          =   195
         Index           =   0
         Left            =   3735
         TabIndex        =   11
         Top             =   405
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cta. Cont:"
         Height          =   195
         Index           =   1
         Left            =   3735
         TabIndex        =   10
         Top             =   765
         Width           =   705
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Item:"
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
         TabIndex        =   9
         Tag             =   "Hilado :"
         Top             =   735
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Presentacion:"
         Height          =   195
         Index           =   2
         Left            =   3705
         TabIndex        =   8
         Top             =   1065
         Width           =   975
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
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
         Index           =   2
         Left            =   195
         TabIndex        =   7
         Tag             =   "Hilado :"
         Top             =   1305
         Width           =   510
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2430
      TabIndex        =   20
      Top             =   5355
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmMantAlmacen.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   555
      Left            =   6690
      TabIndex        =   23
      Top             =   5340
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      Custom          =   "0~0~MOVIMIENTOS~True~True~&Movimientos~0~0~1~Movimientos Permitidos~0~False~False~&Movimientos~Movimientos Permitidos"
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmMantAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Estado As String
Dim Reg As New ADODB.Recordset
Sub Datos(Accion As String, EsAccion As Boolean)

Set Reg = Nothing
Reg.CursorLocation = adUseClient
Reg.Open "UP_Lg_Almacen '" & Accion & "','" & TxtCodigo & "','" & Me.TxtAlmacen & "','" & _
            Trim(Right(CmbTipItem, 2)) & "','" & Me.TxtCuenta & "','" & IIf(optSi, "S", "N") & "','" & Right(cmbPresentacion, 1) & "'", cConnect
            
Set Me.DGridLista.DataSource = Reg
If Not EsAccion Then
    DGridLista_RowColChange 0, 0
End If

End Sub


Sub Habilita()
Me.TxtAlmacen.Enabled = True
Me.TxtCodigo.Enabled = True
Me.TxtCuenta.Enabled = True
Me.cmbPresentacion.Enabled = True
Me.CmbTipItem.Enabled = True
Frame2.Enabled = True
End Sub
Sub Deshabilita()
Me.TxtAlmacen.Enabled = False
Me.TxtCodigo.Enabled = False
Me.TxtCuenta.Enabled = False
Me.cmbPresentacion.Enabled = False
Me.CmbTipItem.Enabled = False
Frame2.Enabled = False
End Sub

Sub Limpia()
Me.TxtAlmacen = ""
Me.TxtCodigo = ""
Me.TxtCuenta = ""
Me.cmbPresentacion.ListIndex = -1
Me.CmbTipItem.ListIndex = -1
End Sub

Private Sub cmdFirst_Click()
    If Not Reg.BOF Then Reg.MoveFirst
End Sub

Private Sub cmdLast_Click()
    If Not Reg.EOF Then Reg.MoveLast
End Sub

Private Sub cmdNext_Click()
If Not Reg.EOF Then Reg.MoveNext
End Sub

Private Sub cmdPrevious_Click()
If Not Reg.BOF Then Reg.MovePrevious
End Sub


Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo hand
If Not Reg.EOF And Not Reg.BOF Then
        TxtCodigo = Reg("cod. almacen")
        TxtAlmacen = Reg("almacen")
        TxtCuenta = Reg("cta. cont")
        BuscaCombo Reg("item"), 1, CmbTipItem
        BuscaCombo Reg("presentacion"), 1, cmbPresentacion
        If Reg("status") = "Operativo" Then
            optSi.Value = True
        Else
            optNo.Value = True
        End If
End If
Exit Sub
hand:
ErrorHandler Err, "DGridLista_RowColChange"
End Sub


Private Sub Form_Load()
'cCONNECT = "Provider=sqloledb;Server=SERVIDOR;Database=lives;UID=sa;pwd=;"
'cSEGURIDAD = "Provider=sqloledb;Server=servidor;Database=seguridad;UID=sa;pwd=;"

LlenaCombo CmbTipItem, "select Des_TipItem +space(100)+Tip_Item from lg_tipitem order by 1", cConnect
LlenaCombo cmbPresentacion, "select Des_Presentacion  +space(100)+Tip_Presentacion  from lg_presentacion order by 1", cConnect

Limpia
Deshabilita
Datos "V", False
MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp1, Me.Name)
FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp1, Me.Name)
Call FormSet(Me)
FormateaGrid Me.DGridLista

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "MOVIMIENTOS"
        FrmMantMovPerm.Almacen = Reg("cod. almacen")
        FrmMantMovPerm.Tip_item = Trim(Right(CmbTipItem, 2))
        FrmMantMovPerm.Show 1
End Select
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
        Me.TxtAlmacen.SetFocus
    Case "ELIMINAR"
        Datos "b", True
        Limpia
        Datos "v", False
        Deshabilita
    Case "GRABAR"
        If Me.TxtAlmacen = "" Then MsgBox "Ingrese el nombre del almacen", vbInformation: Exit Sub
        If CmbTipItem = "" Then MsgBox "Seleccione un tipo de item", vbInformation: Exit Sub
        If Trim(Right(CmbTipItem, 2)) = "H" Or Trim(Right(CmbTipItem, 2)) = "T" Then
            If cmbPresentacion = "" Then MsgBox "Seleccione una presentacion", vbInformation: Exit Sub
        End If
        If Estado = "NUEVO" Then
            Datos "i", True
        Else
            Datos "a", True
        End If
        Limpia
        Deshabilita
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Datos "v", False
    Case "DESHACER"
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Limpia
        Datos "v", False
        Deshabilita
    Case "SALIR"
        Unload Me
End Select

Exit Sub
hand:
ErrorHandler Err, "MantFunc1_ActionClick"
End Sub


