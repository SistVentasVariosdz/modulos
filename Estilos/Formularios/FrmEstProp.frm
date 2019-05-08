VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmEstProp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estilos Propios"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   705
      Left            =   585
      TabIndex        =   37
      Top             =   915
      Width           =   5265
      Begin VB.TextBox TxtCliente3 
         Height          =   315
         Left            =   840
         MaxLength       =   4
         TabIndex        =   40
         Top             =   270
         Width           =   735
      End
      Begin VB.TextBox TxtTemporada 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3420
         TabIndex        =   39
         Top             =   270
         Width           =   1485
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   315
         Left            =   1590
         TabIndex        =   38
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Temporada:"
         Height          =   315
         Index           =   5
         Left            =   2430
         TabIndex        =   42
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   41
         Top             =   270
         Width           =   765
      End
   End
   Begin VB.Frame FraCliEst 
      Height          =   705
      Left            =   585
      TabIndex        =   31
      Top             =   915
      Width           =   5265
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   315
         Left            =   1590
         TabIndex        =   36
         Top             =   270
         Width           =   375
      End
      Begin VB.TextBox TxtEstiloCli 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3420
         TabIndex        =   33
         Top             =   270
         Width           =   1485
      End
      Begin VB.TextBox TxtClienteEst 
         Height          =   315
         Left            =   840
         MaxLength       =   4
         TabIndex        =   32
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   35
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Estilo:"
         Height          =   315
         Index           =   1
         Left            =   2430
         TabIndex        =   34
         Top             =   270
         Width           =   945
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4230
      Top             =   3420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.JPG|*.GIF|*.BMP"
   End
   Begin VB.Frame Frame4 
      Height          =   1845
      Left            =   120
      TabIndex        =   22
      Top             =   4680
      Width           =   7395
      Begin VB.CommandButton cmdIcono2 
         Caption         =   "..."
         Height          =   315
         Left            =   6960
         TabIndex        =   29
         Top             =   750
         Width           =   375
      End
      Begin VB.CommandButton cmdIcono1 
         Caption         =   "..."
         Height          =   315
         Left            =   6960
         TabIndex        =   28
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox TxtIcono2 
         Height          =   315
         Left            =   4500
         TabIndex        =   5
         Top             =   750
         Width           =   2415
      End
      Begin VB.TextBox TxtIcono1 
         Height          =   315
         Left            =   4500
         TabIndex        =   4
         Top             =   300
         Width           =   2415
      End
      Begin VB.ComboBox CmdGruTal 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1230
         Width           =   2415
      End
      Begin VB.ComboBox cmbTipPre 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   750
         Width           =   2415
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   315
         Left            =   1140
         TabIndex        =   1
         Top             =   300
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Imagen 2"
         Height          =   315
         Index           =   4
         Left            =   3750
         TabIndex        =   27
         Top             =   750
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Imagen 1"
         Height          =   315
         Index           =   3
         Left            =   3750
         TabIndex        =   26
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Grupo Talla"
         Height          =   315
         Index           =   2
         Left            =   150
         TabIndex        =   25
         Top             =   1260
         Width           =   1425
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Prenda"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion"
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   23
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2595
      Left            =   120
      TabIndex        =   20
      Top             =   1950
      Width           =   7365
      Begin MSDataGridLib.DataGrid Dg 
         Height          =   2235
         Left            =   90
         TabIndex        =   21
         Top             =   210
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   3942
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "codigo"
            Caption         =   "codigo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Descripcion"
            Caption         =   "Descripcion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Tipo de Prenda"
            Caption         =   "Tipo de Prenda"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Grupo de Talla"
            Caption         =   "Grupo de Talla"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "icono1"
            Caption         =   "Imagen 1"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "icono2"
            Caption         =   "Imagen2"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   810
      TabIndex        =   14
      Top             =   6450
      Width           =   6165
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   120
         Picture         =   "FrmEstProp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Primero"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   600
         Picture         =   "FrmEstProp.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Anterior"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   1110
         Picture         =   "FrmEstProp.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Siguiente"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1590
         Picture         =   "FrmEstProp.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Ultimo"
         Top             =   240
         Width           =   495
      End
      Begin Mantenimientos.MantFunc MantFunc2 
         Height          =   540
         Left            =   2280
         TabIndex        =   19
         Top             =   180
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   953
         Custom          =   $"FrmEstProp.frx":05C8
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
      End
   End
   Begin VB.Frame FraEstilo 
      Height          =   705
      Left            =   570
      TabIndex        =   9
      Top             =   930
      Width           =   5265
      Begin VB.TextBox TxtEstilo 
         Height          =   315
         Left            =   840
         MaxLength       =   5
         TabIndex        =   11
         Top             =   270
         Width           =   735
      End
      Begin VB.TextBox TxtDesEstilo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3420
         TabIndex        =   10
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Descripcion:"
         Height          =   315
         Index           =   0
         Left            =   2430
         TabIndex        =   13
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Estilo:"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   270
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   7245
      Begin VB.OptionButton optTemporada 
         Caption         =   "Cliente / Temporada"
         Height          =   405
         Left            =   5160
         TabIndex        =   8
         Top             =   210
         Width           =   2055
      End
      Begin VB.OptionButton OptCliente 
         Caption         =   "Cliente / Estilo"
         Height          =   405
         Left            =   2520
         TabIndex        =   7
         Top             =   210
         Width           =   1755
      End
      Begin VB.OptionButton optEstilo 
         Caption         =   "Estilo"
         Height          =   405
         Left            =   420
         TabIndex        =   6
         Top             =   210
         Value           =   -1  'True
         Width           =   1365
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   495
      Left            =   6060
      TabIndex        =   30
      Top             =   1080
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
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   75
      Left            =   0
      Top             =   1800
      Width           =   7455
   End
End
Attribute VB_Name = "FrmEstProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo As String
Public Descripcion As String
Dim Estado As String
Dim Reg As New ADODB.Recordset


Dim TmpCliente As String
Dim TmpEstilo As String
Dim TmpTemp As String
Sub Cargar_Datos(pAccion As String, pCliente As String, pEstilo As String, _
                pTemporada As String, pEstiloProp As String, pDescripcion As String, _
                pTipoPrenda As String, pGrupo As String, pIcono1 As String, pIcono2 As String)

On Error GoTo hand
Set Reg = Nothing

Reg.CursorLocation = adUseClient
Reg.Open "UP_Es_EsPro '" & pAccion & "','" & pCliente & "','" & pEstilo & "','" & pTemporada & "','" & pEstiloProp & "','" & pDescripcion & "','" & pTipoPrenda & "','" & pGrupo & "','" & pIcono1 & "','" & pIcono2 & "'", cCONNECT

'Dg.Columns(0).Visible = False
Set Dg.DataSource = Reg
If (pAccion <> "B") Then
    If (pAccion <> "I") Then
        If (pAccion <> "A") Then If Reg.RecordCount > 0 Then Dg_RowColChange 0, 0
    End If
End If
Exit Sub
hand:
ErrorHandler Err, "CARGAR_DATOS"
End Sub


Sub Habilita(modo As Boolean)
Me.TxtDescripcion.Enabled = modo
Me.cmbTipPre.Enabled = modo
Me.CmdGruTal.Enabled = modo
Me.TxtIcono1.Enabled = modo
Me.TxtIcono2.Enabled = modo
Me.cmdIcono1.Enabled = modo
Me.cmdIcono2.Enabled = modo
End Sub

Sub Limpia()
Me.TxtDescripcion = ""
Me.cmbTipPre.ListIndex = -1
Me.CmdGruTal.ListIndex = -1
Me.TxtIcono1 = ""
Me.TxtIcono2 = ""
End Sub

Private Sub cmdFirst_Click()
    If Not Reg.BOF Then Reg.MoveFirst
End Sub

Private Sub cmdIcono1_Click()
With cd
'    .FileName = "XX"
    .ShowSave
    'Me.TxtIcono1 = Mid(.FileName, 1, Len(.FileName) - 2)
    Me.TxtIcono1 = .FileName
End With
End Sub


Private Sub cmdIcono2_Click()
With cd
'    .FileName = "XX"
    .ShowSave
    'Me.TxtIcono2 = Mid(.FileName, 1, Len(.FileName) - 2)
    Me.TxtIcono2 = .FileName
End With
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

Private Sub Command1_Click()
Set frmBusqGeneral.oParent = Me
frmBusqGeneral.sQuery = "Select abr_cliente as Codigo,nom_cliente as Descripcion from tg_cliente order by 1"
frmBusqGeneral.Cargar_Datos

frmBusqGeneral.Show 1
TxtClienteEst = Codigo
TmpCliente = DevuelveCampo("select cod_cliente from tg_cliente where abr_cliente='" & TxtClienteEst & "'", cCONNECT)

End Sub

Private Sub Command2_Click()
Set frmBusqGeneral.oParent = Me
frmBusqGeneral.sQuery = "Select abr_cliente as Codigo,nom_cliente as Descripcion from tg_cliente order by 1"
frmBusqGeneral.Cargar_Datos

frmBusqGeneral.Show 1
TxtCliente3 = Codigo
TmpCliente = DevuelveCampo("select cod_cliente from tg_cliente where abr_cliente='" & TxtCliente3 & "'", cCONNECT)

End Sub


Private Sub Dg_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Not Reg.EOF Then
    Me.TxtDescripcion = Dg.Columns(1).Text
    BuscaCombo Dg.Columns(2).Text, 1, cmbTipPre
    BuscaCombo Dg.Columns(3).Text, 1, CmdGruTal
    
    Me.TxtIcono1 = Dg.Columns(4).Text
    Me.TxtIcono2 = Dg.Columns(5).Text
End If
End Sub


Private Sub Form_Load()
'cCONNECT = "Provider=sqloledb;Server=SERVIDOR;Database=lives;UID=sa;pwd=;"
'cSEGURIDAD = "Provider=sqloledb;Server=servidor;Database=seguridad;UID=sa;pwd=;"

LlenaCombo Me.cmbTipPre, "Select Des_TipPre  + space(100) +Cod_TipPre from tg_tippre order by Des_TipPre  ", cCONNECT
LlenaCombo Me.CmdGruTal, "Select  des_grutal + space(100)+ cod_grutal  from es_tallas  order by des_grutal ", cCONNECT

Cargar_Datos "1", "", "", "", "", "", "", "", "", ""
MantFunc2.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
FormateaGrid Dg
Habilita False
optEstilo_Click
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
If Me.OptCliente Then
    Cargar_Datos "2", TmpCliente, TmpEstilo, TmpTemp, "", "", "", "", "", ""
ElseIf Me.optEstilo Then
    Cargar_Datos "1", TmpCliente, TmpEstilo, TmpTemp, "", "", "", "", "", ""
ElseIf Me.optTemporada Then
    Cargar_Datos "3", TmpCliente, TmpEstilo, TmpTemp, "", "", "", "", "", ""
End If
End Sub


Private Sub MantFunc2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "ADICIONAR"
        HabilitaMant Me.MantFunc2, "GRABAR/DESHACER"
        Limpia
        Habilita True
        Estado = "NUEVO"
    Case "MODIFICAR"
        HabilitaMant Me.MantFunc2, "GRABAR/DESHACER"
        Estado = "MODIFICAR"
        Habilita True
        TxtDescripcion.SetFocus
    Case "ELIMINAR"
        Cargar_Datos "B", "", "", "", Dg.Columns(0).Text, "", _
         "", "", "", ""

            If Me.optEstilo Then
                Cargar_Datos "1", TmpCliente, TmpEstilo, TmpTemp, "", "", "", "", "", ""
            End If
            Limpia
            Habilita False
        
    Case "GRABAR"
            If Trim(TxtDescripcion) = "" Then MsgBox "Llene la descripcion", vbInformation: Exit Sub
            If Me.cmbTipPre = "" Then MsgBox "Llene el tipo de prenda", vbInformation: Exit Sub
            If Me.CmdGruTal = "" Then MsgBox "Llene el grupo de talla", vbInformation: Exit Sub
            
            If Estado = "NUEVO" Then
                Cargar_Datos "I", TmpCliente, TmpEstilo, TmpTemp, "", Me.TxtDescripcion, _
                 Right(Me.cmbTipPre, 3), Right(Me.CmdGruTal, 3), Me.TxtIcono1, TxtIcono2
            Else
                Cargar_Datos "A", TmpCliente, TmpEstilo, TmpTemp, Dg.Columns(0).Text, Me.TxtDescripcion, _
                 Right(Me.cmbTipPre, 3), Right(Me.CmdGruTal, 3), Me.TxtIcono1, TxtIcono2
            End If
            Limpia
            Habilita False
            If Me.optEstilo Then
                Cargar_Datos "1", TmpCliente, TmpEstilo, TmpTemp, "", "", "", "", "", ""
            End If
            HabilitaMant Me.MantFunc2, "ADICIONAR/MODIFICAR/ELIMINAR"
    Case "DESHACER"
        HabilitaMant Me.MantFunc2, "ADICIONAR/MODIFICAR/ELIMINAR"
        
    Case "SALIR"
        Unload Me
End Select

End Sub


Private Sub optcliente_Click()
FraEstilo.Visible = False
FraCliEst.Visible = True
Frame5.Visible = False
End Sub

Private Sub optEstilo_Click()
FraEstilo.Visible = True
FraCliEst.Visible = False
Frame5.Visible = False
End Sub

Private Sub optTemporada_Click()
FraEstilo.Visible = False
FraCliEst.Visible = False
Frame5.Visible = True

End Sub

Private Sub TxtCliente3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ExisteCampo("abr_cliente", "tg_cliente", TxtCliente3, cCONNECT, True) = False Then MsgBox "El cliente no existe", vbInformation: Exit Sub
End If

End Sub


Private Sub TxtClienteEst_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ExisteCampo("abr_cliente", "tg_cliente", TxtClienteEst, cCONNECT, True) = False Then MsgBox "El cliente no existe", vbInformation: Exit Sub
End If
End Sub


Private Sub TxtDesEstilo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Trim(TxtDesEstilo)) < 5 Then MsgBox "Debe tener 5 caracteres de longitud", vbInformation: Exit Sub
    'If ExisteCampo("Des_estpro", "es_estpro", TxtDesEstilo, cCONNECT) Then
        If DevuelveCampo("select count(Cod_EstPro) from es_estpro where Des_estpro like '" & TxtDesEstilo & "%'", cCONNECT) > 1 Then
            frmBusqGeneral.sQuery = "select Cod_EstPro as codigo, Des_estpro as descripcion from es_estpro "
            Set frmBusqGeneral.oParent = Me
            frmBusqGeneral.Cargar_Datos
            frmBusqGeneral.Show 1
            TxtEstilo = Codigo
            TxtDesEstilo = Descripcion
            TmpEstilo = Codigo
        Else
            TxtEstilo = DevuelveCampo("select Cod_EstPro from es_estpro where Des_estpro like '" & TxtDesEstilo & "%'", cCONNECT)
        End If
    'Else
    '    MsgBox "Descripcion no existente", vbInformation
    'End If
End If
End Sub


Private Sub TxtEstilo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ExisteCampo("Cod_EstPro", "es_estpro", Format(TxtEstilo, "#00000"), cCONNECT) Then
        TxtDesEstilo = DevuelveCampo("select Des_estpro from es_estpro where Cod_EstPro='" & Format(TxtEstilo, "#00000") & "'", cCONNECT)
        TxtEstilo = Format(TxtEstilo, "#00000")
        TmpEstilo = Format(TxtEstilo, "#00000")
    Else
        MsgBox "Codigo no existente", vbInformation
    End If
Else
    SoloNumeros TxtEstilo, KeyAscii, False, 0, 5
End If
End Sub


Private Sub TxtEstiloCli_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 '   If ExisteCampo("Des_EstCli", "tg_estclitem", TxtEstiloCli, cCONNECT, True) = False Then MsgBox "El estilo no existe", vbInformation: Exit Sub
        
    If DevuelveCampo("select count(Des_EstCli) from tg_estclitem where Des_EstCli like '" & TxtEstiloCli & "%' and cod_cliente='" & TmpCliente & "'", cCONNECT) > 1 Then
        frmBusqGeneral.sQuery = "select Cod_EstCli as Codigo, Des_EstCli as Descripcion from tg_estclitem order by 2 "
        Set frmBusqGeneral.oParent = Me
        frmBusqGeneral.Cargar_Datos
        frmBusqGeneral.Show 1
        TxtEstiloCli = Descripcion
        TmpEstilo = Codigo
        'TmpCliente = DevuelveCampo("select a.cod_cliente from tg_cliente a, tg_estclitem b where a.cod_cliente=b.cod_cliente and b.Cod_EstCli='" & TmpEstilo & "' and b.Des_EstCli='" & TxtEstiloCli & "'", cCONNECT)
        'TxtClienteEst = DevuelveCampo("select a.abr_cliente from tg_cliente a, tg_estclitem b where a.cod_cliente=b.cod_cliente and b.Cod_EstCli='" & TmpEstilo & "' and b.Des_EstCli='" & TxtEstiloCli & "'", cCONNECT)
    Else
        TxtEstiloCli = DevuelveCampo("select Cod_EstCli from tg_estclitem where Des_EstCli like '" & TxtEstiloCli & "%' and cod_cliente='" & TmpCliente & "'", cCONNECT)
    End If
        
End If
End Sub


'
Private Sub txttemporada_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 '   If ExisteCampo("Des_EstCli", "tg_estclitem", TxtEstiloCli, cCONNECT, True) = False Then MsgBox "El estilo no existe", vbInformation: Exit Sub
        
    If DevuelveCampo("select count(*) from tg_temcli where Nom_TemCli like '" & txttemporada & "%' and cod_cliente='" & TmpCliente & "'", cCONNECT) > 1 Then
        frmBusqGeneral.sQuery = "select Cod_TemCli as Codigo, Nom_TemCli as Descripcion from tg_temcli where Nom_TemCli like '" & txttemporada & "%' and cod_cliente='" & TmpCliente & "' order by 2 "
        Set frmBusqGeneral.oParent = Me
        frmBusqGeneral.Cargar_Datos
        frmBusqGeneral.Show 1
        txttemporada = Descripcion
        TmpTemp = Codigo
        'TmpCliente = DevuelveCampo("select a.cod_cliente from tg_cliente a, tg_estclitem b where a.cod_cliente=b.cod_cliente and b.Cod_EstCli='" & TmpEstilo & "' and b.Des_EstCli='" & TxtEstiloCli & "'", cCONNECT)
        'TxtClienteEst = DevuelveCampo("select a.abr_cliente from tg_cliente a, tg_estclitem b where a.cod_cliente=b.cod_cliente and b.Cod_EstCli='" & TmpEstilo & "' and b.Des_EstCli='" & TxtEstiloCli & "'", cCONNECT)
    Else
        txttemporada = DevuelveCampo("select Cod_TemCli from tg_temcli where Nom_TemCli like '" & txttemporada & "%' and cod_cliente='" & TmpCliente & "'", cCONNECT)
    End If
        
End If

End Sub


