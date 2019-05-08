VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Begin VB.Form FrmMantMovPerm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos Permitidos"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6825
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
      Height          =   1020
      Left            =   60
      TabIndex        =   7
      Tag             =   "Detail"
      Top             =   3360
      Width           =   6660
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
         Left            =   4320
         TabIndex        =   9
         Top             =   210
         Width           =   2145
         Begin VB.OptionButton optNo 
            Caption         =   "NO"
            Height          =   255
            Left            =   1200
            TabIndex        =   11
            Top             =   240
            Width           =   645
         End
         Begin VB.OptionButton optSi 
            Caption         =   "SI"
            Height          =   255
            Left            =   270
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   465
         End
      End
      Begin VB.ComboBox CmbTipoMov 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   330
         Width           =   2355
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Mov:"
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
         TabIndex        =   12
         Tag             =   "Hilado :"
         Top             =   405
         Width           =   690
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
      Left            =   75
      TabIndex        =   5
      Tag             =   "List"
      Top             =   90
      Width           =   6645
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2775
         Left            =   180
         TabIndex        =   6
         Top             =   345
         Width           =   6345
         _ExtentX        =   11192
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   667
      TabIndex        =   0
      Top             =   4500
      Width           =   1965
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "FrmMantMovPerm.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Anterior"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "FrmMantMovPerm.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Siguiente"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "FrmMantMovPerm.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Primero"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "FrmMantMovPerm.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ultimo"
         Top             =   0
         Width           =   495
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2670
      TabIndex        =   13
      Top             =   4470
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmMantMovPerm.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "FrmMantMovPerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Almacen As String
Public Tip_item As String
Dim Estado As String
Dim Reg As New ADODB.Recordset
Sub Datos(Accion As String, EsAccion As Boolean)

Set Reg = Nothing
Reg.CursorLocation = adUseClient
Reg.Open "UP_lg_tipmovialm '" & Accion & "','" & Almacen & "','" & Trim(Right(Me.CmbTipoMov, 4)) & "','" & IIf(Me.optNo, "N", "S") & "'", cConnect

Set Me.DGridLista.DataSource = Reg

Me.DGridLista.Columns(0).Visible = False
If Not EsAccion Then DGridLista_RowColChange 0, 0
End Sub

Sub Deshabilita()
Frame2.Enabled = False
Me.CmbTipoMov.Enabled = False
End Sub

Sub Habilita()
Frame2.Enabled = True
Me.CmbTipoMov.Enabled = True
End Sub

Sub Limpia()
Me.CmbTipoMov.ListIndex = -1

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
If Not Reg.EOF And Not Reg.BOF Then
    BuscaCombo Reg("Tipo Mov."), 1, CmbTipoMov
    
    If Reg("Operativo") = "SI" Then
        Me.optSi.Value = True
    Else
        Me.optNo.Value = True
    End If
    
End If
End Sub

Private Sub Form_Load()

'MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp1, Me.Name)
LlenaCombo Me.CmbTipoMov, "select Des_TipMov + space(100)+ Cod_TipMov from lg_tiposmov where tip_item='" & Tip_item & "' order by 1", cConnect

FormateaGrid Me.DGridLista
Limpia
Datos "V", False
Deshabilita
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand
Select Case ActionName
    Case "ADICIONAR"
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Limpia
        Habilita
        Estado = "NUEVO"
        CmbTipoMov.SetFocus
    Case "MODIFICAR"
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Estado = "MODIFICAR"
        Deshabilita
        Frame2.Enabled = True
        
    Case "ELIMINAR"
        Datos "b", True
        Limpia
        Datos "v", False
        Deshabilita
    Case "GRABAR"
        If Estado = "NUEVO" Then
            If Me.CmbTipoMov = "" Then MsgBox "Seleccione un tipo de movimiento", vbInformation: Exit Sub

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


