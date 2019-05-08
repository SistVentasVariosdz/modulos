VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form FrmManteOrigenImportacion 
   Caption         =   "Origen Importación"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   1200
      TabIndex        =   3
      Top             =   4080
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmManteOrigenImportacion.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin VB.Frame Frame2 
      Height          =   2715
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6135
      Begin GridEX20.GridEX gexOrigenImportacion 
         Height          =   2355
         Left            =   105
         TabIndex        =   0
         Top             =   240
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   4154
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
         FormatStyle(1)  =   "FrmManteOrigenImportacion.frx":0160
         FormatStyle(2)  =   "FrmManteOrigenImportacion.frx":0298
         FormatStyle(3)  =   "FrmManteOrigenImportacion.frx":0348
         FormatStyle(4)  =   "FrmManteOrigenImportacion.frx":03FC
         FormatStyle(5)  =   "FrmManteOrigenImportacion.frx":04D4
         FormatStyle(6)  =   "FrmManteOrigenImportacion.frx":058C
         FormatStyle(7)  =   "FrmManteOrigenImportacion.frx":066C
         ImageCount      =   0
         PrinterProperties=   "FrmManteOrigenImportacion.frx":068C
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   2760
      Width           =   6135
      Begin VB.TextBox TxtDesOrigenImportacion 
         Height          =   330
         Left            =   1920
         TabIndex        =   2
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox TxtCodOrigenImportacion 
         Height          =   330
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Des. Origen Importación"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Origen Importación"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1710
      End
   End
End
Attribute VB_Name = "FrmManteOrigenImportacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Strsql As String
Dim vTipo As String
Dim vMessage As Variant

Private Sub Form_Load()
CARGA_GRID
End Sub

Sub SALVAR_DATOS()
On Error GoTo ErrSlavarDatos

Strsql = "UP_MAN_LG_ORIGEN_IMPORTACION '" & vTipo & "','" & TxtCodOrigenImportacion.Text & "','" & TxtDesOrigenImportacion.Text & "'"
ExecuteSQL cConnect, Strsql

Exit Sub
ErrSlavarDatos:
    ErrorHandler Err, "SALVAR_DATOS"
End Sub

Function VALIDA_ENTRADAS() As Boolean
    If Trim(Me.TxtCodOrigenImportacion.Text) = "" Then
        MsgBox "Ingrese Codigo de Origen Importación", vbInformation
        VALIDA_ENTRADAS = False
        TxtCodOrigenImportacion.SetFocus
        Exit Function
    End If
    If Trim(Me.TxtDesOrigenImportacion.Text) = "" Then
        MsgBox "Ingrese Descripción Origen Importación", vbInformation
        VALIDA_ENTRADAS = False
        TxtDesOrigenImportacion.SetFocus
        Exit Function
    End If
    VALIDA_ENTRADAS = True
End Function

Private Sub gexOrigenImportacion_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
If gexOrigenImportacion.RowCount > 0 Then
    TxtCodOrigenImportacion.Text = gexOrigenImportacion.Value(gexOrigenImportacion.Columns("Cod_OrigenImportacion").Index)
    TxtDesOrigenImportacion.Text = gexOrigenImportacion.Value(gexOrigenImportacion.Columns("Des_OrigenImportacion").Index)
Else
    TxtCodOrigenImportacion.Text = ""
    TxtDesOrigenImportacion.Text = ""
End If
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ADICIONAR"
            vTipo = "I"
            LIMPIA_DATOS
            HABILITA
'            txtSecuencia.Enabled = True
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Case "MODIFICAR"
            If gexOrigenImportacion.RowCount = 0 Then Exit Sub
            vTipo = "U"
            HABILITA
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Case "ELIMINAR"
            If gexOrigenImportacion.RowCount = 0 Then Exit Sub
            vMessage = MsgBox("Esta seguro que desea eliminar el registro", vbYesNo, "Eliminar")
            If vMessage = vbYes Then
                vTipo = "D"
                SALVAR_DATOS
            End If
            CARGA_GRID
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Case "GRABAR"
                If VALIDA_ENTRADAS = False Then Exit Sub
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                SALVAR_DATOS
                CARGA_GRID
                vTipo = ""
        Case "DESHACER"
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            CARGA_GRID
            DESHABILITA
        Case "SALIR"
            FrmImportaciones.TxtCod_OrigenImportacion.Text = Me.TxtCodOrigenImportacion.Text
            FrmImportaciones.TxtDes_OrigenImportacion.Text = Me.TxtDesOrigenImportacion.Text
            Unload Me
    End Select
End Sub

Sub CARGA_GRID()
On Error GoTo ErrCargaGrid
    Strsql = "UP_MAN_LG_ORIGEN_IMPORTACION 'V','',''"
    Set gexOrigenImportacion.ADORecordset = CargarRecordSetDesconectado(Strsql, cConnect)
        
    DESHABILITA
    Exit Sub
ErrCargaGrid:
ErrorHandler Err, "Carga_Grid"
End Sub

Sub LIMPIA_DATOS()
TxtCodOrigenImportacion.Text = ""
TxtDesOrigenImportacion.Text = ""
End Sub

Sub DESHABILITA()
gexOrigenImportacion.Enabled = True
TxtCodOrigenImportacion.Enabled = False
TxtDesOrigenImportacion.Enabled = False
End Sub

Sub HABILITA()
gexOrigenImportacion.Enabled = False
TxtCodOrigenImportacion.Enabled = True
TxtDesOrigenImportacion.Enabled = True
End Sub

Private Sub TxtCodOrigenImportacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtDesOrigenImportacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
