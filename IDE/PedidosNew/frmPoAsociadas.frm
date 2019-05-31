VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmPoAsociadas 
   Caption         =   "PO Asociadas"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   1710
      Left            =   7800
      TabIndex        =   1
      Top             =   480
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   3016
      Custom          =   $"frmPoAsociadas.frx":0000
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX grdPoAsociadas 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7858
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmPoAsociadas.frx":00DE
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmPoAsociadas.frx":0430
      Column(2)       =   "frmPoAsociadas.frx":04F8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmPoAsociadas.frx":059C
      FormatStyle(2)  =   "frmPoAsociadas.frx":06D4
      FormatStyle(3)  =   "frmPoAsociadas.frx":0784
      FormatStyle(4)  =   "frmPoAsociadas.frx":0838
      FormatStyle(5)  =   "frmPoAsociadas.frx":0910
      FormatStyle(6)  =   "frmPoAsociadas.frx":09C8
      FormatStyle(7)  =   "frmPoAsociadas.frx":0AA8
      FormatStyle(8)  =   "frmPoAsociadas.frx":0F60
      ImageCount      =   1
      ImagePicture(1) =   "frmPoAsociadas.frx":13AC
      PrinterProperties=   "frmPoAsociadas.frx":16FE
   End
End
Attribute VB_Name = "frmPoAsociadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public COD_CLIENTE       As String

Public cod_purord        As String

Public cod_lotpurord     As String

Public cod_estcli        As String

Public cod_purodHija     As String

Public cod_lotpurordHija As String

Public cod_estcliHija    As String

Dim strSql               As String

Dim fnuevo               As Boolean

Private Sub Form_Load()
    BUSCAR
End Sub

Sub BUSCAR()

    On Error GoTo drpDepurar

    Dim strSql As String

    strSql = "Tg_Muestra_POS_Hijas '" & COD_CLIENTE & "','" & cod_purord & "','" & cod_lotpurord & "','" & cod_estcli & "'"
    Set grdPoAsociadas.ADORecordset = CargarRecordSetDesconectado(strSql, cCONNECT)
  
    grdPoAsociadas.Columns("Cod_Purord").Width = 1800
    grdPoAsociadas.Columns("Cod_Lotpurord").Width = 1000
    grdPoAsociadas.Columns("Cod_Estcli").Width = 1200
    grdPoAsociadas.Columns("Fec_DespachoAct").Width = 1500
    grdPoAsociadas.Columns("Num_PreReq").Width = 1600
    grdPoAsociadas.Columns("clave").Width = 0
    '  grdPoAsociadas.Columns("TOTAL").Width = 1300
    '
    grdPoAsociadas.Columns("Cod_Purord").Caption = "Purchase Order"
    grdPoAsociadas.Columns("Cod_Lotpurord").Caption = "Estilo Lote"
    grdPoAsociadas.Columns("Cod_Estcli").Caption = "Estilo Numero"
    grdPoAsociadas.Columns("Fec_DespachoAct").Caption = "Despacho Actual"
    grdPoAsociadas.Columns("Num_PreReq").Caption = "Prendas Requeridas"
    '  grdPoAsociadas.Columns("TOTAL").Caption = "TOTAL"
  
    '  grdPoAsociadas.Columns("Fec_Evento").Width = 2115
  
    'lbNroTrabajadores.Caption = grdPoAsociadas.RowCount
  
    Exit Sub

drpDepurar:
    ErrorHandler Err, "Buscar"
  
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, _
                                   ByVal ActionType As Integer, _
                                   ByVal ActionName As String)

    Select Case ActionName

        Case "ASIGNAR"

            Dim fAph As New frmAsignarPo

            Set fAph.oParent = Me
            fAph.Show 1
            BUSCAR

            If cod_purodHija <> "" And cod_lotpurordHija <> "" And cod_estcliHija <> "" Then
                fnuevo = grdPoAsociadas.Find(grdPoAsociadas.Columns("clave").Index, jgexGreaterThanOrEqualTo, COD_CLIENTE + cod_purodHija + cod_lotpurordHija + cod_estcliHija)
            End If

        Case "DESASIGNAR"
            eliminar
            BUSCAR

        Case "SALIR"
            Unload Me
    End Select
  
End Sub

Sub eliminar()

    Dim sSQl As String

    Dim iRet As String

    If grdPoAsociadas.RowCount = 0 Then Exit Sub

    On Error GoTo eliminar

    If MsgBox("Esta Seguro de Eliminar este PO", vbYesNo + vbInformation, "IMPORTANTE") = vbYes Then
        strSql = "Tg_Desasigna_Po_Hija_Madre '" & COD_CLIENTE & "','" & cod_purord & "','" & cod_lotpurord & "','" & cod_estcli & "','" & grdPoAsociadas.value(grdPoAsociadas.Columns("Cod_Purord").Index) & "','" & grdPoAsociadas.value(grdPoAsociadas.Columns("Cod_Lotpurord").Index) & "','" & grdPoAsociadas.value(grdPoAsociadas.Columns("Cod_Estcli").Index) & "'"
        ExecuteCommandSQL cCONNECT, strSql
    
        '    sSQL = "SELECT count(*) FROM Tg_Operario_Productos_Comedor where Cod_Fabrica='002' and Tip_Trabajador='E' and Cod_Trabajador='0503'"
        '    iRet = DevuelveCampo(sSQL, cConnect)
        '    If iRet = 1 Then
        '        sSQL = "SELECT Cod_Fabrica FROM TG_Fabrica "
        '        objFabrica.Text = DevuelveCampo(sSQL, cConnect)
        '
        '        sSQL = "SELECT Nom_Fabrica FROM TG_Fabrica "
        '        objNombreFabrica.Text = DevuelveCampo(sSQL, cConnect)
        '        objFabrica.Enabled = False
        '        objNombreFabrica.Enabled = False
        '
        '    End If
        BUSCAR
    End If

    Exit Sub

eliminar:
    ErrorHandler Err, "Eliminar"
End Sub
