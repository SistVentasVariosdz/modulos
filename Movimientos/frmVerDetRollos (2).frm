VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form frmVerDetRollos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   9180
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4905
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   9090
      Begin GridEX20.GridEX gexMovDet 
         Height          =   4485
         Left            =   135
         TabIndex        =   2
         Top             =   240
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   7911
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ContScroll      =   -1  'True
         AllowEdit       =   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmVerDetRollos.frx":0000
         Column(2)       =   "frmVerDetRollos.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmVerDetRollos.frx":016C
         FormatStyle(2)  =   "frmVerDetRollos.frx":02A4
         FormatStyle(3)  =   "frmVerDetRollos.frx":0354
         FormatStyle(4)  =   "frmVerDetRollos.frx":0408
         FormatStyle(5)  =   "frmVerDetRollos.frx":04E0
         FormatStyle(6)  =   "frmVerDetRollos.frx":0598
         ImageCount      =   0
         PrinterProperties=   "frmVerDetRollos.frx":0678
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   1545
         Left            =   7725
         TabIndex        =   1
         Top             =   270
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   2725
         Custom          =   $"frmVerDetRollos.frx":0850
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   25
      End
   End
End
Attribute VB_Name = "frmVerDetRollos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Almacen As String, sNum_MovStk As String, sCod_TipMov As String
Dim StrSql As String

Dim sFlg_Devolucion_Rollos_Tejeduria As String

Public Sub BUSCAR()
On Error GoTo Fin
Dim sTit As String
    sTit = "Mostrar Movimentos Stock Rollos"
    
    StrSql = "EXEC Tj_SM_MUESTRA_MOV_TELA_CRUDA_ROLLOS '" & sCod_Almacen & "', '" & _
             sNum_MovStk & "'"
    Set gexMovDet.ADORecordset = CargarRecordSetDesconectado(StrSql, cCONNECT)
    
    gexMovDet.Columns("Cod_Almacen").Visible = False
    gexMovDet.Columns("Num_MovStk").Visible = False
    gexMovDet.Columns("Num_Secuencia").Caption = "Sec."
    gexMovDet.Columns("Cod_OrdTra").Caption = "O.T."
    gexMovDet.Columns("Num_Secuencia_OT").Visible = False
    gexMovDet.Columns("Num_Rollo").Visible = False
    gexMovDet.Columns("Prefijo_Maquina").Caption = "P.Maq."
    gexMovDet.Columns("Codigo_Rollo").Caption = "Cod.Rollo"
    gexMovDet.Columns("Kgs_Rollo").Caption = "Kgs.Rollo"
    gexMovDet.Columns("Uni_Rollos").Caption = "Uni.Rollos"
    gexMovDet.Columns("Cod_Calidad").Visible = False
    gexMovDet.Columns("Des_Calidad").Caption = "Calidad"
    gexMovDet.Columns("Observacion").Caption = "Observaciones"
    gexMovDet.Columns("Cod_TipMov").Visible = False
    
    gexMovDet.Columns("Cod_Almacen").Width = 1500
    gexMovDet.Columns("Num_MovStk").Width = 1500
    gexMovDet.Columns("Num_Secuencia").Width = 465
    gexMovDet.Columns("Cod_OrdTra").Width = 525
    gexMovDet.Columns("Num_Secuencia_OT").Width = 1500
    gexMovDet.Columns("Prefijo_Maquina").Width = 700
    gexMovDet.Columns("Codigo_Rollo").Width = 900
    gexMovDet.Columns("Kgs_Rollo").Width = 810
    gexMovDet.Columns("Uni_Rollos").Width = 855
    gexMovDet.Columns("Cod_Calidad").Width = 1500
    gexMovDet.Columns("Des_Calidad").Width = 1290
    gexMovDet.Columns("Observacion").Width = 2505
    gexMovDet.Columns("Cod_TipMov").Width = 1500
Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, sTit
End Sub

Private Sub Form_Load()
    FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp1, Me.Name)
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "ADICIONAR"
        frmShow_TxMovStk.AddRollo "I", sCod_Almacen, sCod_TipMov, sNum_MovStk
        BUSCAR
    Case "ELIMINAR"
        sFlg_Devolucion_Rollos_Tejeduria = DevuelveCampo("select isnull(Flg_Devolucion_Rollos_Tejeduria,'') from tx_tiposmov where cod_tipmov = '" & sCod_TipMov & "'", cCONNECT)
        If sFlg_Devolucion_Rollos_Tejeduria = "S" Then
            MsgBox "Tipo Movimiento no permite eliminación", vbCritical
            Exit Sub
        End If
        DelDetMov
    Case "SALIR"
        Unload Me
    End Select
End Sub

Private Sub DelDetMov()
On Error GoTo Fin
Dim sTit As String
    If gexMovDet.RowCount = 0 Then Exit Sub
    sTit = "Eliminar Detalle de Movimiento"
    If MsgBox("Desea Eliminar este Movimento?", vbQuestion + vbYesNo, sTit) = vbNo Then Exit Sub
    
    StrSql = "EXEC LG_UP_MAN_TX_MOVISTK_DETALLE_ROLLOS 'D', '" & _
    gexMovDet.Value(gexMovDet.Columns("Cod_Almacen").Index) & "', '" & _
    gexMovDet.Value(gexMovDet.Columns("Num_MovStk").Index) & "', '" & _
    gexMovDet.Value(gexMovDet.Columns("Num_Secuencia").Index) & "', '" & _
    gexMovDet.Value(gexMovDet.Columns("Prefijo_Maquina").Index) & "', '" & _
    gexMovDet.Value(gexMovDet.Columns("Codigo_Rollo").Index) & "', " & _
    gexMovDet.Value(gexMovDet.Columns("Kgs_Rollo").Index) & ", " & _
    gexMovDet.Value(gexMovDet.Columns("Uni_Rollos").Index) & ", '" & _
    gexMovDet.Value(gexMovDet.Columns("observacion").Index) & "'"
    
    ExecuteSQL cCONNECT, StrSql
    
    BUSCAR
Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, sTit
End Sub
