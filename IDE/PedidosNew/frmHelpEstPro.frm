VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmHelpEstPro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estilos Propios"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   9105
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3285
      TabIndex        =   4
      Top             =   3675
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmHelpEstPro.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   3525
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   8970
      Begin GridEX20.GridEX gexList 
         Height          =   2685
         Left            =   90
         TabIndex        =   3
         Top             =   765
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   4736
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmHelpEstPro.frx":008F
         Column(2)       =   "frmHelpEstPro.frx":0157
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmHelpEstPro.frx":01FB
         FormatStyle(2)  =   "frmHelpEstPro.frx":0333
         FormatStyle(3)  =   "frmHelpEstPro.frx":03E3
         FormatStyle(4)  =   "frmHelpEstPro.frx":0497
         FormatStyle(5)  =   "frmHelpEstPro.frx":056F
         FormatStyle(6)  =   "frmHelpEstPro.frx":0627
         ImageCount      =   0
         PrinterProperties=   "frmHelpEstPro.frx":0707
      End
      Begin VB.TextBox TxtEstCli 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1530
         TabIndex        =   1
         Top             =   285
         Width           =   3585
      End
      Begin VB.Label Label1 
         Caption         =   "Estilo Cliente:"
         Height          =   240
         Left            =   195
         TabIndex        =   2
         Top             =   345
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmHelpEstPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sCod_Cliente   As String

Public sCod_TemCli    As String

Public sCod_PurOrd    As String

Public sCod_LotPurOrd As String

Dim strSql            As String

Sub CARGA_GRID()

    Dim oGroup As GridEX20.JSGroup

    On Error GoTo hand

    strSql = "exec sm_ayuda_estilos_propios_Por_Estilo_Cliente '" & sCod_Cliente & "','" & sCod_TemCli & "','" & TxtEstCli.Text & "'"

    gexList.ClearFields

    Set gexList.ADORecordset = CargarRecordSetDesconectado(strSql, cCONNECT)

    gexList.Columns("Estilo").Visible = False

    gexList.Columns("version").Width = 1000

    Set oGroup = gexList.Groups.Add(gexList.Columns("Estilo").Index, jgexSortAscending)

    gexList.DefaultGroupMode = jgexDGMCollapsed
    gexList.CollapseAll

    Exit Sub

hand:
    ErrorHandler Err, "CARGA_GRID"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, _
                                   ByVal ActionType As Integer, _
                                   ByVal ActionName As String)

    Select Case ActionName

        Case "ACEPTAR"

            If gexList.RowCount = 0 Then Exit Sub
            If gexList.GetRowData(gexList.Row).RowType = jgexRowTypeRecord Then
                SALVAR_DATOS_A_TEMPORAL
                Unload Me
            End If

        Case "CANCELAR"
            Unload Me
    End Select

End Sub

Sub SALVAR_DATOS_A_TEMPORAL()

    On Error GoTo errsalvar

    strSql = "EXEC SM_TM_LOTESTPRO '" & vusu & "','" & sCod_Cliente & "','" & sCod_PurOrd & "','" & Me.TxtEstCli.Text & "','" & gexList.value(gexList.Columns("COD_ESTPRO").Index) & "'," & gexList.value(gexList.Columns("NUM_SOLICITUD_CONS").Index) & ",'" & gexList.value(gexList.Columns("COD_ESTPRO").Index) & "','" & gexList.value(gexList.Columns("VERSION").Index) & "'"
            
    Call ExecuteCommandSQL(cCONNECT, strSql)

    Exit Sub

errsalvar:
    ErrorHandler Err, "SALVAR_DATOS"
End Sub

Private Sub gexList_DblClick()
    Call FunctButt1_ActionClick(1, 0, "ACEPTAR")
End Sub

Private Sub gexList_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call gexList_DblClick
    End If

End Sub
