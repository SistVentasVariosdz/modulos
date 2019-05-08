VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form frmShowTX_Rapport 
   Caption         =   "Rapport"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   10245
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   5460
      Left            =   45
      TabIndex        =   7
      Top             =   1125
      Width           =   10125
      Begin GridEX20.GridEX GridEX1 
         Height          =   5130
         Left            =   90
         TabIndex        =   8
         Top             =   210
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   9049
         Version         =   "2.0"
         HoldSortSettings=   -1  'True
         DefaultGroupMode=   1
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmShowTX_Rapport.frx":0000
         Column(2)       =   "frmShowTX_Rapport.frx":00C8
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmShowTX_Rapport.frx":016C
         FormatStyle(2)  =   "frmShowTX_Rapport.frx":02A4
         FormatStyle(3)  =   "frmShowTX_Rapport.frx":0354
         FormatStyle(4)  =   "frmShowTX_Rapport.frx":0408
         FormatStyle(5)  =   "frmShowTX_Rapport.frx":04E0
         FormatStyle(6)  =   "frmShowTX_Rapport.frx":0598
         FormatStyle(7)  =   "frmShowTX_Rapport.frx":0678
         FormatStyle(8)  =   "frmShowTX_Rapport.frx":0724
         ImageCount      =   0
         PrinterProperties=   "frmShowTX_Rapport.frx":07D4
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Argumentos de Búsqueda"
      Height          =   1155
      Left            =   45
      TabIndex        =   0
      Top             =   -30
      Width           =   10140
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   510
         Left            =   8715
         TabIndex        =   3
         Top             =   315
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   900
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.TextBox TxtRapport 
         Height          =   315
         Left            =   1755
         TabIndex        =   2
         Top             =   270
         Width           =   1185
      End
      Begin VB.TextBox TxtClienteEst 
         Height          =   315
         Left            =   2805
         TabIndex        =   10
         Top             =   675
         Width           =   2115
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   330
         Left            =   2385
         TabIndex        =   9
         Top             =   660
         Width           =   375
      End
      Begin VB.TextBox txtCodCliEst 
         Height          =   315
         Left            =   1770
         TabIndex        =   1
         Top             =   660
         Width           =   555
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   210
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "RN"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   330
         Width           =   1740
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1320
      TabIndex        =   4
      Top             =   6600
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   900
      Custom          =   $"frmShowTX_Rapport.frx":09AC
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1170
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmShowTX_Rapport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public tmpCliente As String
Public codigo As String
Public descripcion As String
Public sCod_Cliente As String
Dim StrSql  As String

Dim Opcion As Integer
Dim rslista As ADODB.Recordset

Dim mensaje As String

Private Sub Form_Load()
Opcion = 2

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "ADICIONAR"
        Load frmTX_Rapport
        frmTX_Rapport.Opcion = "I"
        frmTX_Rapport.Rapport = 0
        If Option2.Value = True Then
            If txtCodCliEst.Text <> "" Then
                frmTX_Rapport.TxtCod_Cliente.Text = txtCodCliEst
            End If
        End If
        frmTX_Rapport.Show 1
        If frmTX_Rapport.Rapport <> 0 Then
            Option1.Value = True
            TxtRapport.Text = frmTX_Rapport.Rapport
            Call CargaRapport
        End If
    Case "MODIFICAR"
        Load frmTX_Rapport
        frmTX_Rapport.Opcion = "U"
        frmTX_Rapport.Rapport = GridEX1.Value(GridEX1.Columns("rapport_number").Index)
        frmTX_Rapport.txtDesRapport = GridEX1.Value(GridEX1.Columns("DESCRIPCION").Index)
        frmTX_Rapport.txtcod_tela = GridEX1.Value(GridEX1.Columns("cod_tela").Index)
        frmTX_Rapport.txtdes_tela = GridEX1.Value(GridEX1.Columns("DES_tela").Index)
        frmTX_Rapport.TxtCod_Cliente = GridEX1.Value(GridEX1.Columns("abr_cliente").Index)
        frmTX_Rapport.TxtDes_Cliente = GridEX1.Value(GridEX1.Columns("nom_cliente").Index)
        frmTX_Rapport.Show 1
        Option1.Value = True
        Me.TxtRapport.Text = frmTX_Rapport.Rapport
        Call CargaRapport
    Case "ELIMINAR"
        If GridEX1.RowCount > 0 Then
            mensaje = MsgBox("¿Confirma que desea eliminar el registro?", vbYesNo)
            If mensaje = vbYes Then
                Call ELIMINAR_RAPPORT
            End If
            Call CargaRapport
        End If
    Case "COMBINACIONES"
        If GridEX1.RowCount > 0 Then
            Load frmShowTX_Rapport_Comb
            frmShowTX_Rapport_Comb.TxtRapport = GridEX1.Value(GridEX1.Columns("rapport_number").Index)
            frmShowTX_Rapport_Comb.Tela = GridEX1.Value(GridEX1.Columns("cod_tela").Index)
            frmShowTX_Rapport_Comb.CARGA_GRID
            frmShowTX_Rapport_Comb.Show 1
        Else
            MsgBox ("No existen registros")
        End If
    Case "COMPOSICION"
        If GridEX1.RowCount > 0 Then
            Load frmShowTX_Rapport_Composicion
            frmShowTX_Rapport_Composicion.TxtRapport = GridEX1.Value(GridEX1.Columns("rapport_number").Index)
            frmShowTX_Rapport_Composicion.CARGA_GRID
            frmShowTX_Rapport_Composicion.Show 1
        Else
            MsgBox ("No existen registros")
        End If
    Case "SALIR"
        Unload Me
End Select
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Call CargaRapport
End Sub

Private Sub Option1_Click()
Opcion = 1
End Sub

Private Sub Option2_Click()
    Opcion = 2
End Sub

Private Sub txtCodCliEst_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ExisteCampo("abr_cliente", "tg_cliente", txtCodCliEst.Text, cCONNECT, True) = False Then
        MsgBox "El cliente no existe", vbInformation
        Exit Sub
    Else
        If Trim(txtCodCliEst.Text) = "" Then
            Command1_Click
        Else
            tmpCliente = DevuelveCampo("Select cod_cliente from tg_cliente  where abr_cliente='" & txtCodCliEst.Text & "'", cCONNECT)
            StrSql = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE cod_cliente='" & tmpCliente & "'"
            TxtClienteEst.Text = DevuelveCampo(StrSql, cCONNECT)
            
        End If
        FunctButt2.SetFocus
    End If
End If

End Sub

Private Sub Command1_Click()
Set frmBusqGeneral.oParent = Me
frmBusqGeneral.sQuery = "Select abr_cliente as Codigo,nom_cliente as Descripcion from tg_cliente order by 1"
frmBusqGeneral.Cargar_Datos

frmBusqGeneral.Show 1
TxtClienteEst = descripcion
txtCodCliEst.Text = codigo
tmpCliente = DevuelveCampo("Select cod_cliente from tg_cliente  where abr_cliente='" & txtCodCliEst.Text & "'", cCONNECT)
sCod_Cliente = tmpCliente

End Sub

Private Sub TxtClienteEst_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ExisteCampo("abr_cliente", "tg_cliente", TxtClienteEst, cCONNECT, True) = False Then
        MsgBox "El cliente no existe", vbInformation
        Exit Sub
    Else
    End If
End If
End Sub

Private Sub CargaRapport()
    StrSql = "EXEC UP_SEL_RAPPORT " & Opcion & ",'" & TxtRapport.Text & "','" & tmpCliente & "'"
    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(StrSql, cCONNECT)
    GridEX1.Columns("cod_usuario").Visible = False
    GridEX1.Columns("fec_ultmod").Visible = False
    GridEX1.Columns("cod_estacion").Visible = False
    GridEX1.Columns("rapport_number").Caption = "Rapport"
    GridEX1.Columns("rapport_number").Width = 700
    GridEX1.Columns("descripcion").Caption = "Desc. Rapport"
    GridEX1.Columns("descripcion").Width = 1500
    GridEX1.Columns("Cod_tela").Caption = "Cod. Tela"
    GridEX1.Columns("cod_tela").Width = 1000
    GridEX1.Columns("DES_tela").Caption = "Desc. Tela"
    GridEX1.Columns("DES_tela").Width = 3200
    GridEX1.Columns("ABR_cliente").Caption = "Abr. Cliente"
    GridEX1.Columns("ABR_cliente").Width = 900
    GridEX1.Columns("NOM_cliente").Width = 2500
    
End Sub

Sub ELIMINAR_RAPPORT()
    Dim con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim StrSql As String
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    con.ConnectionString = cCONNECT
    con.Open
    
    con.BeginTrans

    StrSql = "EXEC UP_MAN_TX_RAPPORT 'D'," & GridEX1.Value(GridEX1.Columns("RAPPORT_NUMBER").Index) & ",'','','','','',''"
                
    con.Execute StrSql
    con.CommitTrans
    
    Screen.MousePointer = vbDefault
    MsgBox "Rapport eliminado ", vbInformation, "Mensaje"
    Exit Sub
    
Salvar_DatosErr:
    con.RollbackTrans
    Set con = Nothing
    Screen.MousePointer = vbDefault
    ErrorHandler Err, "ELIMINAR_RAPPORT"
End Sub

Private Sub TxtRapport_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FunctButt2.SetFocus
End If
End Sub
