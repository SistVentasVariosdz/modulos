VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmPresentaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presentaciones"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   135
      TabIndex        =   5
      Top             =   4830
      Width           =   1965
      Begin VB.CommandButton cmdClose 
         Height          =   495
         Left            =   4575
         Picture         =   "FrmPresentaciones.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Cerrar"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNew 
         Height          =   495
         Left            =   2160
         Picture         =   "FrmPresentaciones.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Nuevo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdEdit 
         Height          =   495
         Left            =   2640
         Picture         =   "FrmPresentaciones.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Editar"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   495
         Left            =   3120
         Picture         =   "FrmPresentaciones.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Eliminar"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   3600
         Picture         =   "FrmPresentaciones.frx":05C8
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Grabar"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdUndo 
         Height          =   495
         Left            =   4080
         Picture         =   "FrmPresentaciones.frx":073A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Deshacerundo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "FrmPresentaciones.frx":08AC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "FrmPresentaciones.frx":0A1E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "FrmPresentaciones.frx":0B90
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "FrmPresentaciones.frx":0D02
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
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
      Height          =   1050
      Left            =   135
      TabIndex        =   2
      Tag             =   "Detail"
      Top             =   3645
      Width           =   5655
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
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   3
         Top             =   390
         Width           =   4365
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion:"
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
         Left            =   135
         TabIndex        =   4
         Tag             =   "Description:"
         Top             =   450
         Width           =   900
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
      Height          =   3450
      Left            =   120
      TabIndex        =   0
      Tag             =   "List"
      Top             =   105
      Width           =   5655
      Begin MSDataGridLib.DataGrid DGridlista 
         Height          =   2970
         Left            =   180
         TabIndex        =   1
         Top             =   345
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   5239
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
            DataField       =   "Descr"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   0
            BeginProperty Column00 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3600
            EndProperty
         EndProperty
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2115
      TabIndex        =   16
      Top             =   4905
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmPresentaciones.frx":0E74
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   45
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmPresentaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Reg As New ADODB.Recordset
Public Codigo As String
Dim Estado As String
Dim TmpPresent As String
Sub Accion(pTipo As String, pPresentacion As String, pDescripcion As String, Optional EsAccion As Boolean = False)
    Set Reg = Nothing
    Reg.CursorLocation = adUseClient
    Reg.Open "UP_Es_EstProPre '" & pTipo & "','" & Codigo & "','" & pPresentacion & "','" & pDescripcion & "'", cCONNECT
If EsAccion = True Then
    
Else
    Set Me.DGridLista.DataSource = Reg
    If Reg.RecordCount > 0 Then DGridLista_RowColChange 0, 0
End If
Me.DGridLista.Columns(0).Visible = False
End Sub


Sub Habilita(modo As Boolean)
Me.TxtDescripcion.Enabled = modo
End Sub

Sub Limpia()
Me.TxtDescripcion = ""
End Sub

Private Sub cmdFirst_Click()
    If Not Reg.BOF Then Reg.MoveFirst
End Sub


Private Sub cmdLast_Click()
    If Not Reg.EOF Then Reg.MoveLast
    
End Sub

Private Sub cmdNext_Click()
If Not Reg.EOF Then Reg.MoveNext
If Reg.EOF Then Reg.MoveLast
End Sub

Private Sub cmdPrevious_Click()
If Not Reg.BOF Then Reg.MovePrevious
If Reg.BOF Then Reg.MoveFirst
End Sub


Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Not Reg.EOF And Not Reg.BOF Then
    Me.TxtDescripcion = Me.DGridLista.Columns(1).Text
    TmpPresent = Me.DGridLista.Columns(0).Text
End If
End Sub


Private Sub Form_Load()
'cCONNECT = "Provider=sqloledb;Server=SERVIDOR;Database=lives;UID=sa;pwd=;"
'cSEGURIDAD = "Provider=sqloledb;Server=servidor;Database=seguridad;UID=sa;pwd=;"
Call FormSet(Me)
MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
Habilita False
FormateaGrid Me.DGridLista
Accion "V", "", "", False
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "ADICIONAR"
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Limpia
        Habilita True
        Estado = "NUEVO"
        Me.TxtDescripcion.SetFocus
    Case "MODIFICAR"
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Me.TxtDescripcion.Enabled = True
        Estado = "MODIFICAR"
    Case "ELIMINAR"
        Accion "B", TmpPresent, "", True
        Limpia
        Habilita False
        Accion "V", " ", "", False
    
    Case "GRABAR"
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        If Estado = "NUEVO" Then
            Accion "I", "", Me.TxtDescripcion, True
        Else
            Accion "A", TmpPresent, Me.TxtDescripcion, True
        End If
        Limpia
        Habilita False
        Accion "V", "", "", False
    Case "DESHACER"
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Habilita False
        Accion "V", "", "", False
    Case "SALIR"
        Unload Me
End Select

End Sub







