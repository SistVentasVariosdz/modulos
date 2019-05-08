VERSION 5.00
Begin VB.Form frmDatosComplementarios 
   Caption         =   "Datos Complementarios"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Grabar"
      Height          =   480
      Left            =   3060
      TabIndex        =   8
      Top             =   3825
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   480
      Left            =   4545
      TabIndex        =   9
      Top             =   3825
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Height          =   3645
      Left            =   105
      TabIndex        =   10
      Top             =   0
      Width           =   8700
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   285
         Left            =   8025
         TabIndex        =   22
         Top             =   300
         Width           =   450
      End
      Begin VB.TextBox txtComposicion 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1695
         Width           =   6060
      End
      Begin VB.TextBox txtDes_Adicional_Partida 
         Height          =   285
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1245
         Width           =   6060
      End
      Begin VB.TextBox txtObservacion 
         Height          =   285
         Left            =   1110
         TabIndex        =   7
         Top             =   3105
         Width           =   6735
      End
      Begin VB.TextBox txtNum_Partida_Arancelaria_Exterior 
         Height          =   285
         Left            =   1110
         TabIndex        =   6
         Top             =   2655
         Width           =   1740
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   285
         Left            =   8010
         TabIndex        =   15
         Top             =   2175
         Width           =   450
      End
      Begin VB.CommandButton cmdPartidaArancelaria 
         Caption         =   "..."
         Height          =   285
         Left            =   8010
         TabIndex        =   14
         Top             =   795
         Width           =   450
      End
      Begin VB.TextBox txtDes_Categoria_Internacional 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   2160
         Width           =   6060
      End
      Begin VB.TextBox txtNum_Categoria_Internacional 
         Height          =   285
         Left            =   1110
         TabIndex        =   4
         Top             =   2160
         Width           =   600
      End
      Begin VB.TextBox txtDes_SecPartida_Arancelaria 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   795
         Width           =   6060
      End
      Begin VB.TextBox txtSec_Partida_Arancelaria 
         Height          =   285
         Left            =   1110
         TabIndex        =   2
         Top             =   795
         Width           =   600
      End
      Begin VB.TextBox txtDes_Partida_Arancelaria 
         Height          =   285
         Left            =   2925
         TabIndex        =   1
         Top             =   315
         Width           =   4950
      End
      Begin VB.TextBox txtNum_Partida_Arancelaria 
         Height          =   285
         Left            =   1095
         TabIndex        =   0
         Top             =   315
         Width           =   1740
      End
      Begin VB.Label Label7 
         Caption         =   "Composición"
         Height          =   300
         Left            =   1065
         TabIndex        =   21
         Top             =   1725
         Width           =   675
      End
      Begin VB.Label Label6 
         Caption         =   "Descrip. Adicional"
         Height          =   465
         Left            =   1095
         TabIndex        =   20
         Top             =   1170
         Width           =   675
      End
      Begin VB.Label Label5 
         Caption         =   "Observac."
         Height          =   285
         Left            =   165
         TabIndex        =   17
         Top             =   3135
         Width           =   1200
      End
      Begin VB.Label Label4 
         Caption         =   "Part H.T.S."
         Height          =   240
         Left            =   165
         TabIndex        =   16
         Top             =   2700
         Width           =   810
      End
      Begin VB.Label Label3 
         Caption         =   "Sec. Partida"
         Height          =   495
         Left            =   150
         TabIndex        =   13
         Top             =   735
         Width           =   840
      End
      Begin VB.Label Label2 
         Caption         =   "Categoria Internacional"
         Height          =   435
         Left            =   150
         TabIndex        =   12
         Top             =   2130
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Partida Arancelaria"
         Height          =   435
         Left            =   165
         TabIndex        =   11
         Top             =   285
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmDatosComplementarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo As String
Public Descripcion As String
Public TipoAdd As String
Public TipoAdd2 As String
Public sCod_Cliente As String
Public sCod_EstCli As String

Dim strSQL  As String

Private Sub CmdAceptar_Click()
    Grabar
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdPartidaArancelaria_Click()
    Load frmMantPartidasArancelarias
    frmMantPartidasArancelarias.sNum_Partida_Arancelaria = txtNum_Partida_Arancelaria.Text
    frmMantPartidasArancelarias.CARGA_GRID
    frmMantPartidasArancelarias.Show vbModal
    Set frmMantPartidasArancelarias = Nothing
End Sub

Private Sub Command2_Click()
    Load frmMantPartArancelariasCab
    frmMantPartArancelariasCab.CARGA_GRID
    frmMantPartArancelariasCab.Show vbModal
    
    Set frmMantPartArancelariasCab = Nothing
End Sub

Private Sub txtNum_Partida_Arancelaria_Exterior_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDes_Partida_Arancelaria_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaPartidaArancelaria 1
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtNum_Categoria_Internacional_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BuscaCategoria 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDES_Categoria_Internacional_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BuscaCategoria 2
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtNum_Partida_Arancelaria_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaPartidaArancelaria 1
        SendKeys "{TAB}"
    End If
End Sub


Public Sub BuscaPartidaArancelaria(Opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Num_Partida_Arancelaria, Des_Partida_Arancelaria FROM TG_PARTIDA_ARANCELARIA WHERE "
    
    txtNum_Partida_Arancelaria = Trim(txtNum_Partida_Arancelaria)
    txtDes_Partida_Arancelaria = Trim(txtDes_Partida_Arancelaria)
    
    Select Case Opcion
    Case 1: strSQL = strSQL & "Num_Partida_Arancelaria like '%" & txtNum_Partida_Arancelaria & "%'"
    Case 2: strSQL = strSQL & "Des_Partida_Arancelaria LIKE '%" & txtDes_Partida_Arancelaria & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.sQuery = strSQL
    frmBusqGeneral3.Cargar_Datos
    Set rstAux = frmBusqGeneral3.DGridLista.ADORecordset
    
    frmBusqGeneral3.DGridLista.Columns("Num_Partida_Arancelaria").Width = 2000
    frmBusqGeneral3.DGridLista.Columns("Des_Partida_Arancelaria").Width = 7000
    
    frmBusqGeneral3.DGridLista.Columns("Num_Partida_Arancelaria").Caption = "Partida"
    frmBusqGeneral3.DGridLista.Columns("Des_Partida_Arancelaria").Caption = "Descrip."
    
    If frmBusqGeneral3.DGridLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtNum_Partida_Arancelaria = ""
    txtDes_Partida_Arancelaria = ""
    
    If Codigo <> "" Then
        txtNum_Partida_Arancelaria = Codigo
        txtDes_Partida_Arancelaria = Descripcion
        
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
End Sub


Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSec_Partida_Arancelaria_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaPartidaArancelaria_Detalle 1
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtDes_SecPartida_Arancelaria_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Public Sub BuscaPartidaArancelaria_Detalle(Opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Sec_Partida_Arancelaria, Des_Partida , Des_Adicional_Partida, Composicion FROM TG_PARTIDA_ARANCELARIA_DETALLE   WHERE Num_Partida_Arancelaria = '" & txtNum_Partida_Arancelaria & "' AND "
    
    txtSec_Partida_Arancelaria = Trim(txtSec_Partida_Arancelaria)
    txtDes_SecPartida_Arancelaria = Trim(txtDes_SecPartida_Arancelaria)
    
    Select Case Opcion
    Case 1: strSQL = strSQL & "Sec_Partida_Arancelaria like '%" & txtSec_Partida_Arancelaria & "%'"
    Case 2: strSQL = strSQL & "Des_Partida  LIKE '%" & txtDes_SecPartida_Arancelaria & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.sQuery = strSQL
    frmBusqGeneral3.Cargar_Datos
    Set rstAux = frmBusqGeneral3.DGridLista.ADORecordset
    
    frmBusqGeneral3.DGridLista.Columns("Sec_Partida_Arancelaria").Width = 500
    frmBusqGeneral3.DGridLista.Columns("Des_Partida").Width = 3000
    frmBusqGeneral3.DGridLista.Columns("Des_Adicional_Partida").Width = 3000
    
    frmBusqGeneral3.DGridLista.Columns("Sec_Partida_Arancelaria").Caption = "Sec.Partida"
    frmBusqGeneral3.DGridLista.Columns("Des_Partida").Caption = "Descrip."
    frmBusqGeneral3.DGridLista.Columns("Des_Adicional_Partida").Caption = "Descrip.Adic"
    
    If frmBusqGeneral3.DGridLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtSec_Partida_Arancelaria = ""
    txtDes_SecPartida_Arancelaria = ""
    
    If Codigo <> "" Then
        txtSec_Partida_Arancelaria = Codigo
        txtDes_SecPartida_Arancelaria = Descripcion
        txtDes_Adicional_Partida = TipoAdd
        txtComposicion = TipoAdd2
    End If
    
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
End Sub


Private Sub Grabar()
On Error GoTo errx

strSQL = "EXEC TG_UPDATE_DATOS_COMPLEM_ESTCLI '" & sCod_Cliente & "','" & sCod_EstCli & "','" & txtNum_Partida_Arancelaria & "','" & txtSec_Partida_Arancelaria & "','" & txtNum_Categoria_Internacional & "','" & txtNum_Partida_Arancelaria_Exterior.Text & "','" & txtObservacion.Text & "'"
ExecuteCommandSQL cCONNECT, strSQL

Unload Me

Exit Sub
errx:
    errores Err.Number
End Sub


Public Sub CargarDatos()
On Error GoTo errx
Dim rs As ADODB.Recordset

Set rs = GetRecordset(cCONNECT, "EXEC TG_MUESTRA_DATOS_COMPLE_ESTCLI '" & sCod_Cliente & "','" & sCod_EstCli & "'")

If Not rs Is Nothing Then
    txtNum_Partida_Arancelaria = RTrim(rs!Num_Partida_Arancelaria)
    txtDes_Partida_Arancelaria = RTrim(rs!des_partida_arancelaria)
    txtSec_Partida_Arancelaria = RTrim(rs!Sec_Partida_Arancelaria)
    txtDes_SecPartida_Arancelaria = RTrim(rs!des_partida)
    txtNum_Categoria_Internacional = RTrim(rs!Num_Categoria_Internacional)
    txtDes_Categoria_Internacional = RTrim(rs!Des_Categoria_Internacional)
    txtDes_Adicional_Partida = RTrim(rs!Des_Adicional_Partida)
    txtComposicion = RTrim(rs!Composicion)
    txtNum_Partida_Arancelaria_Exterior = RTrim(rs!Num_Partida_Arancelaria_Exterior)
    txtObservacion.Text = RTrim(rs!Observacion)
    rs.Close
    
End If
Set rs = Nothing

Exit Sub
errx:
    errores Err.Number
End Sub

Public Sub BuscaCategoria(Opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Num_Categoria_Internacional, Des_Categoria_Internacional FROM TG_Categoria_Estilo WHERE "
    
    txtNum_Categoria_Internacional = Trim(txtNum_Categoria_Internacional)
    txtDes_Categoria_Internacional = Trim(txtDes_Categoria_Internacional)
    
    Select Case Opcion
    Case 1: strSQL = strSQL & "Num_Categoria_Internacional like '%" & txtNum_Categoria_Internacional & "%'"
    Case 2: strSQL = strSQL & "Des_Categoria_Internacional LIKE '%" & txtDes_Categoria_Internacional & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.sQuery = strSQL
    frmBusqGeneral3.Cargar_Datos
    Set rstAux = frmBusqGeneral3.DGridLista.ADORecordset
    
    frmBusqGeneral3.DGridLista.Columns("Num_Categoria_Internacional").Width = 1500
    frmBusqGeneral3.DGridLista.Columns("Des_Categoria_Internacional").Width = 7000
    
    frmBusqGeneral3.DGridLista.Columns("Num_Categoria_Internacional").Caption = "Categoria"
    frmBusqGeneral3.DGridLista.Columns("Des_Categoria_Internacional").Caption = "Descrip."
    
    If frmBusqGeneral3.DGridLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtNum_Categoria_Internacional = ""
    txtDes_Categoria_Internacional = ""
    
    If Codigo <> "" Then
        txtNum_Categoria_Internacional = Codigo
        txtDes_Categoria_Internacional = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
End Sub

