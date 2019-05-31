VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Begin VB.Form FrmComentario 
   Caption         =   "Comentarios"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Comment"
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
      Height          =   3345
      Left            =   120
      TabIndex        =   14
      Tag             =   "List"
      Top             =   60
      Width           =   6705
      Begin MSDataGridLib.DataGrid DGridlista 
         Height          =   2865
         Left            =   90
         TabIndex        =   15
         Top             =   345
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   5054
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
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
               LCID            =   3082
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
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   0
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
      Height          =   1725
      Left            =   135
      TabIndex        =   11
      Tag             =   "Detail"
      Top             =   3495
      Width           =   6705
      Begin VB.TextBox txtComentario 
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
         Height          =   1155
         Left            =   1110
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   12
         Top             =   420
         Width           =   5505
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Comentarios:"
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
         Left            =   90
         TabIndex        =   13
         Tag             =   "Comentaries:"
         Top             =   420
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   675
      TabIndex        =   0
      Top             =   5310
      Width           =   1965
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "FrmComentario.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "FrmComentario.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "FrmComentario.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "FrmComentario.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdUndo 
         Height          =   495
         Left            =   4080
         Picture         =   "FrmComentario.frx":05C8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Deshacerundo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   3600
         Picture         =   "FrmComentario.frx":073A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Grabar"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   495
         Left            =   3120
         Picture         =   "FrmComentario.frx":08AC
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Eliminar"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdEdit 
         Height          =   495
         Left            =   2640
         Picture         =   "FrmComentario.frx":0A1E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Editar"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNew 
         Height          =   495
         Left            =   2160
         Picture         =   "FrmComentario.frx":0B90
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Nuevo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdClose 
         Height          =   495
         Left            =   4575
         Picture         =   "FrmComentario.frx":0D02
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Cerrar"
         Top             =   120
         Width           =   495
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   570
      Left            =   2655
      TabIndex        =   16
      Top             =   5385
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmComentario.frx":0E74
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "FrmComentario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstBusca          As New ADODB.Recordset

Dim Compu             As String

Public Cliente        As String

Public PurOrd         As String

Public Lote           As String

Public EstCli         As String

Public Secuencia      As String

Public oParent        As Object

Public sCod_Cliente   As String

Public sCod_PurOrd    As String

Public sCod_LotPurOrd As String

Public sCod_EstCli    As String

Dim Estado            As String

Public sNivAccUsuario As String

Sub ACTUALIZAR()

    On Error GoTo hand

    B_db.Execute "SM_TG_LotEstCom 'Actualizar','" & Trim(sCod_Cliente) & "','" & Trim(sCod_PurOrd) & "','" & Trim(sCod_EstCli) & "','" & Trim(sCod_LotPurOrd) & "','" & Secuencia & "','" & Trim(txtComentario) & "','" & Format(Date, "dd/mm/yyyy") & "','" & vusu & "','" & Compu & "'"
    Cargar_Data

    Exit Sub

hand:
    ErrorHandler Err, "ACTUALIZAR"
End Sub

Sub Cargar_Data()

    On Error GoTo hand

    Set RstBusca = Nothing
    RstBusca.CursorType = adOpenDynamic
    RstBusca.CursorLocation = adUseClient
    RstBusca.Open "SM_TG_LotEstCom 'ver','" & sCod_Cliente & "','" & sCod_PurOrd & "','" & sCod_EstCli & "','" & sCod_LotPurOrd & "','','','','',''", cCONNECT

    Set DGridlista.DataSource = RstBusca
    'DGridlista.Columns(0).Visible = False
    DGridlista.Columns(0).Visible = True

    Exit Sub

hand:
    ErrorHandler Err, "Cargar_Data"
End Sub

Sub eliminar()

    On Error GoTo hand

    B_db.Execute "SM_TG_LotEstCom 'eliminar','" & Trim(sCod_Cliente) & "','" & Trim(sCod_PurOrd) & "','" & Trim(sCod_EstCli) & "','" & Trim(sCod_LotPurOrd) & "','" & Secuencia & "','','','',''"
    Cargar_Data

    Exit Sub

hand:
    ErrorHandler Err, "ELIMINAR"
End Sub

Sub INSERTAR()

    On Error GoTo hand

    B_db.Execute "SM_TG_LotEstCom 'insertar','" & Trim(sCod_Cliente) & "','" & Trim(sCod_PurOrd) & "','" & Trim(sCod_EstCli) & "','" & Trim(sCod_LotPurOrd) & "','" & Format(DevuelveCampo("SM_TG_LotEstCom 'Clave','','','','','','','','',''", cCONNECT), "0##") & "','" & Trim(txtComentario) & "','" & Format(Date, "dd/mm/yyyy") & "','" & vusu & "','" & Compu & "'"
    Cargar_Data

    Exit Sub

hand:
    ErrorHandler Err, "INSERTAR"
End Sub

Private Sub cmdFirst_Click()

    If Not RstBusca.BOF Then RstBusca.MoveFirst
End Sub

Private Sub cmdLast_Click()

    If Not RstBusca.EOF Then RstBusca.MoveLast
End Sub

Private Sub cmdNext_Click()

    If Not RstBusca.EOF Then RstBusca.MoveNext
End Sub

Private Sub cmdPrevious_Click()

    If Not RstBusca.BOF Then RstBusca.MovePrevious
End Sub

Private Sub DGridlista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If (RstBusca.EOF = False) Then
        If (RstBusca.BOF = False) Then txtComentario = DGridlista.Columns(1).Text: Secuencia = DGridlista.Columns(0).Text
    End If

End Sub

Private Sub Form_Load()

    On Error GoTo hand

    'Me.MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    Call FormSet(Me)
    Set B_db = Nothing
    B_db.ConnectionString = cCONNECT
    B_db.Open

    FormateaGrid DGridlista
    Cargar_Data
    Compu = ComputerName

    Exit Sub

hand:
    ErrorHandler Err, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
    Unload Me
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, _
                                  ByVal ActionType As Integer, _
                                  ByVal ActionName As String)

    On Error GoTo hand:

    Select Case ActionName

        Case "ADICIONAR"
            Estado = "Adicionar"
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            txtComentario.Enabled = True
            txtComentario.SetFocus
            DGridlista.Enabled = False

        Case "MODIFICAR"
            Estado = "Modificar"
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            txtComentario.Enabled = True
            txtComentario.SetFocus

            DGridlista.Enabled = False

        Case "ELIMINAR"

            Dim sTit As String
    
            sTit = "Eliminar Datos Familiares"
    
            If MsgBox("Desea Eliminar Comentario", vbQuestion + vbYesNo, sTit) = vbNo Then Exit Sub

            eliminar
            txtComentario.Enabled = False

        Case "GRABAR"

            If Len(Trim(txtComentario)) = 0 Then MsgBox "Debe ingresar un comentario", vbInformation, "Mensaje":  Exit Sub
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"

            If Estado = "Modificar" Then
                ACTUALIZAR
            Else
                INSERTAR
            End If

            DGridlista.Enabled = True
            txtComentario.Enabled = False

        Case "DESHACER"
            Cargar_Data
            txtComentario.Enabled = False
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridlista.Enabled = True

        Case "SALIR"
            Unload Me
    End Select

    Exit Sub

hand:
    ErrorHandler Err, "MantFunc1_ActionClick"

End Sub

