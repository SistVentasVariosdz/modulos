VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmMantProEst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesos de Estilos"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
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
      Height          =   3855
      Left            =   60
      TabIndex        =   16
      Tag             =   "List"
      Top             =   90
      Width           =   5655
      Begin MSDataGridLib.DataGrid DGridlista 
         Height          =   3375
         Left            =   180
         TabIndex        =   17
         Top             =   345
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
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
            DataField       =   "Codigo"
            Caption         =   "Codigo"
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
            Caption         =   "Descripción"
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
            MarqueeStyle    =   0
            BeginProperty Column00 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3330.142
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
      Height          =   1185
      Left            =   75
      TabIndex        =   13
      Tag             =   "Detail"
      Top             =   4005
      Width           =   5655
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
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
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   1
         Top             =   360
         Width           =   405
      End
      Begin VB.TextBox txtDescripcion 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   2
         Top             =   720
         Width           =   2685
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
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
         Left            =   360
         TabIndex        =   15
         Tag             =   "Code:"
         Top             =   420
         Width           =   945
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
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
         Left            =   330
         TabIndex        =   14
         Tag             =   "Description:"
         Top             =   780
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   75
      TabIndex        =   0
      Top             =   5310
      Width           =   1965
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "FrmMantProEst.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "FrmMantProEst.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "FrmMantProEst.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "FrmMantProEst.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdUndo 
         Height          =   495
         Left            =   4080
         Picture         =   "FrmMantProEst.frx":05C8
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Deshacerundo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   3600
         Picture         =   "FrmMantProEst.frx":073A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Grabar"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   495
         Left            =   3120
         Picture         =   "FrmMantProEst.frx":08AC
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Eliminar"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdEdit 
         Height          =   495
         Left            =   2640
         Picture         =   "FrmMantProEst.frx":0A1E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Editar"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNew 
         Height          =   495
         Left            =   2160
         Picture         =   "FrmMantProEst.frx":0B90
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Nuevo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdClose 
         Height          =   495
         Left            =   4575
         Picture         =   "FrmMantProEst.frx":0D02
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cerrar"
         Top             =   120
         Width           =   495
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   570
      Left            =   2055
      TabIndex        =   18
      Top             =   5385
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmMantProEst.frx":0E74
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "FrmMantProEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstTalla As New ADODB.Recordset
Dim Estado As String
Sub ACTUALIZAR(pCodigo As String, pDescripcion As String)
On Error GoTo hand

    Dim Con As New ADODB.Connection
    Dim Strsql As String
    
    Con.ConnectionString = cCONNECT
    Con.Open
    
    Con.BeginTrans

    Con.Execute "UP_Tg_ProEst 'A','" & pCodigo & "','" & pDescripcion & "'"
    
    Con.CommitTrans
    
    Cargar_Data
Exit Sub
hand:
    Con.RollbackTrans
    ErrorHandler Err, "ACTUALIZAR"
End Sub


Sub Cargar_Data()
On Error GoTo hand
Set RstTalla = Nothing
RstTalla.CursorType = adOpenStatic
RstTalla.CursorLocation = adUseClient
RstTalla.Open "UP_Tg_ProEst 'V','',''", cCONNECT

Set Me.DGridLista.DataSource = Nothing
If RstTalla.RecordCount > 0 Then
    Set Me.DGridLista.DataSource = RstTalla
    DGridlista_RowColChange 0, 0
Else
    Set DGridLista.DataSource = Nothing
    DGridLista.Refresh
    txtCodigo = ""
    TxtDescripcion = ""
End If
Exit Sub
hand:
    ErrorHandler Err, "CARGAR_DATA"
End Sub


Sub ELIMINAR(pCodigo As String)
On Error GoTo hand

    Dim Con As New ADODB.Connection
    Dim Strsql As String
    
    Con.ConnectionString = cCONNECT
    Con.Open

    Con.BeginTrans
        Con.Execute "UP_Tg_ProEst 'B','" & pCodigo & "',''"
    Con.CommitTrans
    
    Cargar_Data
Exit Sub
hand:
Con.RollbackTrans
ErrorHandler Err, "ELIMINAR"
End Sub

Sub INSERTAR(pCodigo As String, pDescripcion As String)
On Error GoTo hand:

    Dim Con As New ADODB.Connection
    Dim Strsql As String
    
    Con.ConnectionString = cCONNECT
    Con.Open
    Con.BeginTrans

    Con.Execute "UP_Tg_ProEst 'I','" & pCodigo & "','" & (pDescripcion) & "'"
    
    Con.CommitTrans
    Cargar_Data
Exit Sub
hand:
Con.RollbackTrans
ErrorHandler Err, "INSERTAR"
End Sub

Private Sub cmdFirst_Click()
    If Not RstTalla.BOF Then RstTalla.MoveFirst
End Sub

Private Sub cmdLast_Click()
    If Not RstTalla.EOF Then RstTalla.MoveLast
End Sub

Private Sub cmdNext_Click()
If Not RstTalla.EOF Then RstTalla.MoveNext
End Sub

Private Sub cmdPrevious_Click()
If Not RstTalla.BOF Then RstTalla.MovePrevious
End Sub

Private Sub DGridlista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If (RstTalla.EOF = False) Then
        If (RstTalla.BOF = False) Then Me.txtCodigo = DGridLista.Columns(0).Text: TxtDescripcion = DGridLista.Columns(1).Text
    End If
End Sub


Private Sub Form_Load()
FormateaGrid Me.DGridLista
Call FormSet(Me)
Me.MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
Cargar_Data

End Sub


Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim varEliminar As Integer
On Error GoTo hand:
Select Case ActionName
    Case "ADICIONAR"
        Estado = "Adicionar"
        TxtDescripcion.Enabled = True
        txtCodigo.Enabled = True
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        TxtDescripcion = ""
        txtCodigo = ""
        txtCodigo.SetFocus
        txtCodigo.BackColor = vbWhite
    Case "MODIFICAR"
        Estado = "Modificar"
        txtCodigo.Enabled = False
        TxtDescripcion.Enabled = True
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        TxtDescripcion.SetFocus
    Case "ELIMINAR"
         varEliminar = MsgBox("Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo + vbDefaultButton2, "Mensaje")
         If varEliminar = vbYes Then
            ELIMINAR txtCodigo
            TxtDescripcion.Enabled = False
        End If
    Case "GRABAR"
            If Len(Trim(TxtDescripcion)) = 0 Or Len(Trim(txtCodigo)) = 0 Then MsgBox "Debe llenar todos los datos", vbInformation, "Mensaje": Exit Sub
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            If Estado = "Modificar" Then
                ACTUALIZAR txtCodigo, TxtDescripcion
            Else
            
                If DevuelveCampo("Select count(*) from tg_proest where Cod_ProEst='" & txtCodigo & "'", cCONNECT) > 0 Then
                    MsgBox "El codigo ya existe", vbInformation
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
                    Exit Sub
                Else
                    INSERTAR txtCodigo, TxtDescripcion
                End If
            End If
            txtCodigo.Enabled = False
            DGridLista.Enabled = True
            TxtDescripcion.Enabled = False
            txtCodigo.BackColor = Deshabilitado
    Case "DESHACER"
        Cargar_Data
        txtCodigo.Enabled = False
        TxtDescripcion.Enabled = False
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        DGridLista.Enabled = True
        txtCodigo.BackColor = Deshabilitado
    Case "SALIR"
        Unload Me
End Select
Exit Sub
hand:
ErrorHandler Err, "MantFunc1_ActionClick"
End Sub







