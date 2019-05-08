VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form FrmManteEmbarques 
   Caption         =   "Embarques"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   6135
      Begin VB.TextBox TxtCodEmbarque 
         Height          =   330
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtDesEmbarque 
         Height          =   330
         Left            =   1920
         TabIndex        =   4
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Embarque"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Des. Embarque"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin GridEX20.GridEX gexEmbarque 
         Height          =   2355
         Left            =   105
         TabIndex        =   1
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
         FormatStyle(1)  =   "FrmManteEmbarques.frx":0000
         FormatStyle(2)  =   "FrmManteEmbarques.frx":0138
         FormatStyle(3)  =   "FrmManteEmbarques.frx":01E8
         FormatStyle(4)  =   "FrmManteEmbarques.frx":029C
         FormatStyle(5)  =   "FrmManteEmbarques.frx":0374
         FormatStyle(6)  =   "FrmManteEmbarques.frx":042C
         FormatStyle(7)  =   "FrmManteEmbarques.frx":050C
         ImageCount      =   0
         PrinterProperties=   "FrmManteEmbarques.frx":052C
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   1200
      TabIndex        =   5
      Top             =   3960
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmManteEmbarques.frx":0704
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "FrmManteEmbarques"
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

Strsql = "UP_MAN_TG_TIPEMB '" & vTipo & "','" & TxtCodEmbarque.Text & "','" & TxtDesEmbarque.Text & "'"
ExecuteSQL cConnect, Strsql

Exit Sub
ErrSlavarDatos:
    ErrorHandler Err, "SALVAR_DATOS"
End Sub

Function VALIDA_ENTRADAS() As Boolean
    If Trim(Me.TxtCodEmbarque.Text) = "" Then
        MsgBox "Ingrese Codigo de Embarque", vbInformation
        VALIDA_ENTRADAS = False
        TxtCodEmbarque.SetFocus
        Exit Function
    End If
    If Trim(Me.TxtDesEmbarque.Text) = "" Then
        MsgBox "Ingrese Descripción DesEmbarque", vbInformation
        VALIDA_ENTRADAS = False
        TxtDesEmbarque.SetFocus
        Exit Function
    End If
    VALIDA_ENTRADAS = True
End Function

Private Sub gexEmbarque_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
If gexEmbarque.RowCount > 0 Then
    TxtCodEmbarque.Text = gexEmbarque.Value(gexEmbarque.Columns("Cod_Embarque").Index)
    TxtDesEmbarque.Text = gexEmbarque.Value(gexEmbarque.Columns("Des_Embarque").Index)
Else
    TxtCodEmbarque.Text = ""
    TxtDesEmbarque.Text = ""
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
            If gexEmbarque.RowCount = 0 Then Exit Sub
            vTipo = "U"
            HABILITA
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Case "ELIMINAR"
            If gexEmbarque.RowCount = 0 Then Exit Sub
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
            FrmImportaciones.TxtCod_Embarque.Text = Me.TxtCodEmbarque.Text
            FrmImportaciones.TxtDes_Embarque.Text = Me.TxtDesEmbarque.Text
            Unload Me
    End Select
End Sub

Sub CARGA_GRID()
On Error GoTo ErrCargaGrid
    Strsql = "UP_MAN_TG_TIPEMB 'V','',''"
    Set gexEmbarque.ADORecordset = CargarRecordSetDesconectado(Strsql, cConnect)
        
    DESHABILITA
    Exit Sub
ErrCargaGrid:
ErrorHandler Err, "Carga_Grid"
End Sub

Sub LIMPIA_DATOS()
TxtCodEmbarque.Text = ""
TxtDesEmbarque.Text = ""
End Sub

Sub DESHABILITA()
gexEmbarque.Enabled = True
TxtCodEmbarque.Enabled = False
TxtDesEmbarque.Enabled = False
End Sub

Sub HABILITA()
gexEmbarque.Enabled = False
TxtCodEmbarque.Enabled = True
TxtDesEmbarque.Enabled = True
End Sub


Private Sub TxtCodEmbarque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtDesEmbarque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
