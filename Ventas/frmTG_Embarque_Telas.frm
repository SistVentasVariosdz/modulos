VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmTG_Embarque_Telas 
   Caption         =   "Detalle Embarque Telas"
   ClientHeight    =   5256
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   5256
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Detalle Real"
      Height          =   2652
      Left            =   4440
      TabIndex        =   23
      Top             =   1920
      Width           =   3012
      Begin VB.TextBox txtUnidadesReal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   37
         Tag             =   "SET"
         Text            =   "0"
         Top             =   2160
         Width           =   1200
      End
      Begin VB.TextBox txtPeso_Bruto_Real 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   28
         Tag             =   "SET"
         Text            =   "0"
         Top             =   360
         Width           =   1200
      End
      Begin VB.TextBox txtPeso_Neto_Real 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   27
         Tag             =   "SET"
         Text            =   "0"
         Top             =   720
         Width           =   1200
      End
      Begin VB.TextBox txtRollosReal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   26
         Tag             =   "SET"
         Text            =   "0"
         Top             =   1080
         Width           =   1200
      End
      Begin VB.TextBox txtUbicajeReal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   25
         Tag             =   "SET"
         Text            =   "0"
         Top             =   1440
         Width           =   1200
      End
      Begin VB.TextBox txtKgsReal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   24
         Tag             =   "SET"
         Text            =   "0"
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Label Label16 
         Caption         =   "Unidades  :"
         Height          =   252
         Left            =   120
         TabIndex        =   36
         Top             =   2160
         Width           =   732
      End
      Begin VB.Label Label14 
         Caption         =   "Peso Bruto :"
         Height          =   252
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label13 
         Caption         =   "Peso Neto :"
         Height          =   252
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   972
      End
      Begin VB.Label Label12 
         Caption         =   "Rollos :"
         Height          =   252
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   972
      End
      Begin VB.Label Label11 
         Caption         =   "CUbicaje :"
         Height          =   252
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   972
      End
      Begin VB.Label Label10 
         Caption         =   "Kgs :"
         Height          =   252
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   972
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle Programado"
      Height          =   2652
      Left            =   360
      TabIndex        =   12
      Top             =   1920
      Width           =   3132
      Begin VB.TextBox txtUnidadesProg 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1680
         TabIndex        =   35
         Tag             =   "SET"
         Text            =   "0"
         Top             =   2160
         Width           =   1200
      End
      Begin VB.TextBox txtKgsProg 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1680
         TabIndex        =   22
         Tag             =   "SET"
         Text            =   "0"
         Top             =   1800
         Width           =   1200
      End
      Begin VB.TextBox txtUbicajeProg 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1680
         TabIndex        =   20
         Tag             =   "SET"
         Text            =   "0"
         Top             =   1440
         Width           =   1200
      End
      Begin VB.TextBox txtRollosProg 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1680
         TabIndex        =   18
         Tag             =   "SET"
         Text            =   "0"
         Top             =   1080
         Width           =   1200
      End
      Begin VB.TextBox txtPeso_Neto_Prog 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1680
         TabIndex        =   16
         Tag             =   "SET"
         Text            =   "0"
         Top             =   720
         Width           =   1200
      End
      Begin VB.TextBox txtPeso_Bruto_Prog 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1680
         TabIndex        =   14
         Tag             =   "SET"
         Text            =   "0"
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label15 
         Caption         =   "Unidades  :"
         Height          =   252
         Left            =   120
         TabIndex        =   34
         Top             =   2160
         Width           =   732
      End
      Begin VB.Label Label9 
         Caption         =   "Kgs :"
         Height          =   252
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   732
      End
      Begin VB.Label Label8 
         Caption         =   "CUbicaje :"
         Height          =   252
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   972
      End
      Begin VB.Label Label7 
         Caption         =   "Rollos :"
         Height          =   252
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   972
      End
      Begin VB.Label Label6 
         Caption         =   "Peso Neto :"
         Height          =   252
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   972
      End
      Begin VB.Label Label5 
         Caption         =   "Peso Bruto :"
         Height          =   252
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   972
      End
   End
   Begin VB.TextBox txtDesUniMedida 
      Height          =   288
      Left            =   3000
      TabIndex        =   11
      Top             =   1440
      Width           =   4452
   End
   Begin VB.TextBox txtCodUniMedida 
      Height          =   288
      Left            =   1920
      TabIndex        =   10
      Top             =   1440
      Width           =   1212
   End
   Begin VB.TextBox txtDesColor 
      Height          =   288
      Left            =   3000
      TabIndex        =   8
      Top             =   1080
      Width           =   4452
   End
   Begin VB.TextBox txtCodColor 
      Height          =   288
      Left            =   1920
      TabIndex        =   7
      Top             =   1080
      Width           =   1212
   End
   Begin VB.TextBox txtDesComb 
      Height          =   288
      Left            =   3000
      TabIndex        =   5
      Top             =   600
      Width           =   4452
   End
   Begin VB.TextBox txtDesTela 
      Height          =   288
      Left            =   3000
      TabIndex        =   4
      Top             =   200
      Width           =   4452
   End
   Begin VB.TextBox txtCodComb 
      Height          =   288
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   1212
   End
   Begin VB.TextBox txtCodTela 
      Height          =   288
      Left            =   1920
      TabIndex        =   1
      Top             =   200
      Width           =   1212
   End
   Begin FunctionsButtons.FunctButt FunctOKCancel 
      Height          =   516
      Left            =   2520
      TabIndex        =   38
      Top             =   4680
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   910
      Custom          =   $"frmTG_Embarque_Telas.frx":0000
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Label Label4 
      Caption         =   "Unidad Medida"
      Height          =   372
      Left            =   360
      TabIndex        =   9
      Top             =   1440
      Width           =   1332
   End
   Begin VB.Label Label3 
      Caption         =   "Color."
      Height          =   252
      Left            =   360
      TabIndex        =   6
      Top             =   1116
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "Comb."
      Height          =   252
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Tela"
      Height          =   252
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "frmTG_Embarque_Telas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lNum_Embarque As Long
Public codigo As String
Public Descripcion As String
Public Saccion As String
Public lSec_Embarque As Integer
Public oParent As Object
Dim strSQL As String


Private Sub Form_Load()
    Me.txtCodUniMedida = "KG"
    'Me.txtDesUniMedida = "KILOS"
End Sub

Private Sub FunctOKCancel_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            GrabaEmbarqueTelas
        Case "CANCELAR"
            Unload Me
    End Select
End Sub

Private Sub txtCodColor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.txtCodColor.Text) = "" Then
            Call Me.BuscaColor(3)
        Else
            Call Me.BuscaColor(1)
        End If
    End If
End Sub

Private Sub txtCodComb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.txtCodComb.Text) = "" Then
            Call Me.BuscaComb(3)
        Else
            Call Me.BuscaComb(1)
        End If
    End If
End Sub

Private Sub txtCodTela_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.txtCodTela.Text) = "" Then
            Call Me.BuscaTela(3)
        Else
            Call Me.BuscaTela(1)
        End If
    End If
End Sub

Private Sub txtDesColor_eyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.txtDesColor.Text) = "" Then
            Call Me.BuscaColor(3)
        Else
            Call Me.BuscaColor(2)
        End If
    End If
End Sub

Private Sub txtCodUniMedida_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.txtCodUniMedida.Text) = "" Then
            Call Me.BuscaUnidadMedida(3)
        Else
            Call Me.BuscaUnidadMedida(1)
        End If
    End If
End Sub

Private Sub txtDesComb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.txtDesComb.Text) = "" Then
            Call Me.BuscaComb(3)
        Else
            Call Me.BuscaComb(2)
        End If
    End If
End Sub
Private Sub txtDesTela_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      If Trim(txtDesTela.Text) = "" Then
        Call Me.BuscaTela(3)
      Else
         Call Me.BuscaTela(2)
      End If
    End If
End Sub

Private Sub txtDesUniMedida_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.txtCodUniMedida.Text) = "" Then
            Call Me.BuscaUnidadMedida(3)
        Else
            Call Me.BuscaUnidadMedida(2)
        End If
    End If
End Sub


Public Sub BuscaTela(opcion As Integer)
Dim sField As String, iRows As Long
Dim rstAux As ADODB.Recordset
    
    strSQL = " SELECT RTRIM(Cod_Tela) as 'Codigo' , RTRIM(Des_Tela) AS 'DescripciÛn' FROM tx_tela WHERE "
    txtCodTela = Trim(txtCodTela)
    txtDesTela = Trim(txtDesTela)

    Select Case opcion
    Case 1: strSQL = strSQL & " Cod_Tela   like '%" & Trim(txtCodTela.Text) & "%'  "
    Case 2: strSQL = strSQL & " Des_Tela  like '%" & Trim(txtDesTela.Text) & "%' "
    Case 3: strSQL = " SELECT RTRIM(Cod_Tela) as 'Codigo' , RTRIM(Des_Tela) AS 'DescripciÛn' FROM tx_tela  "
    End Select
    
    
    txtCodTela = ""
    txtDesTela = ""
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
        .Cargar_Datos
        
        codigo = ""
        Descripcion = ""
        
        iRows = .DGridLista.RowCount
        Set rstAux = .DGridLista.ADORecordset
        If .DGridLista.RowCount > 1 Then
            .Show vbModal
        ElseIf .DGridLista.RowCount = 1 Then
            codigo = .DGridLista.Value(.DGridLista.Columns("CODIGO").Index)
            Descripcion = .DGridLista.Value(.DGridLista.Columns("DESCRIPCIÛN").Index)
        End If
        
        If codigo <> "" Then
            txtCodTela = RTrim(codigo)
            txtDesTela = RTrim(Descripcion)
            txtCodComb.SetFocus
        End If
            
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub

Public Sub BuscaComb(opcion As Integer)
Dim sField As String, iRows As Long
Dim rstAux As ADODB.Recordset
    
    strSQL = " SELECT RTRIM(Cod_Comb) as 'Codigo' , RTRIM(Des_Comb) AS 'DescripciÛn' FROM tx_telaComb WHERE cod_tela='" & Me.txtCodTela & "' and "
    txtCodComb = Trim(txtCodComb)
    txtDesComb = Trim(txtDesComb)

    Select Case opcion
    Case 1: strSQL = strSQL & " Cod_Comb   like '%" & Trim(txtCodComb.Text) & "%'  "
    Case 2: strSQL = strSQL & " Des_Comb  like '%" & Trim(txtDesComb.Text) & "%' "
    Case 3: strSQL = " SELECT RTRIM(Cod_Comb) as 'Codigo' , RTRIM(Des_Comb) AS 'DescripciÛn' FROM tx_telaComb WHERE cod_tela='" & Me.txtCodTela & "'"
    End Select
        
        
    txtCodComb = ""
    txtDesComb = ""
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
        .Cargar_Datos
        
        codigo = ""
        Descripcion = ""
        
        iRows = .DGridLista.RowCount
        Set rstAux = .DGridLista.ADORecordset
        If .DGridLista.RowCount > 1 Then
            .Show vbModal
        ElseIf .DGridLista.RowCount = 1 Then
            codigo = .DGridLista.Value(.DGridLista.Columns("CODIGO").Index)
            Descripcion = .DGridLista.Value(.DGridLista.Columns("DESCRIPCIÛN").Index)
        End If
        
        If codigo <> "" Then
            txtCodComb = RTrim(codigo)
            txtDesComb = RTrim(Descripcion)
            txtCodColor.SetFocus
        End If
            
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub

Public Sub BuscaUnidadMedida(opcion As Integer)
Dim sField As String, iRows As Long
Dim rstAux As ADODB.Recordset
    
    strSQL = " SELECT RTRIM(Cod_UniMed) as 'Codigo' , RTRIM(Des_UniMed) AS 'DescripciÛn' FROM tg_uniMed where "
    txtCodUniMedida = Trim(txtCodUniMedida)
    txtDesUniMedida = Trim(txtDesUniMedida)

    Select Case opcion
    Case 1: strSQL = strSQL & " Cod_UniMed   like '%" & Trim(txtCodUniMedida.Text) & "%'  "
    Case 2: strSQL = strSQL & " Des_UniMed  like '%" & Trim(txtDesUniMedida.Text) & "%' "
    Case 3: strSQL = " SELECT RTRIM(Cod_UniMed) as 'Codigo' , RTRIM(Des_UniMed) AS 'DescripciÛn' FROM tg_uniMed"
    End Select
    
    
    txtCodUniMedida = ""
    txtDesUniMedida = ""
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
        .Cargar_Datos
        
        codigo = ""
        Descripcion = ""
        
        iRows = .DGridLista.RowCount
        Set rstAux = .DGridLista.ADORecordset
        If .DGridLista.RowCount > 1 Then
            .Show vbModal
        ElseIf .DGridLista.RowCount = 1 Then
            codigo = .DGridLista.Value(.DGridLista.Columns("CODIGO").Index)
            Descripcion = .DGridLista.Value(.DGridLista.Columns("DESCRIPCIÛN").Index)
        End If
        
        If codigo <> "" Then
            txtCodUniMedida = RTrim(codigo)
            txtDesUniMedida = RTrim(Descripcion)
            Me.txtPeso_Bruto_Prog.SetFocus
        End If
            
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub

Public Sub BuscaColor(opcion As Integer)
Dim sField As String, iRows As Long
Dim rstAux As ADODB.Recordset
    
    strSQL = " SELECT RTRIM(Cod_Color) as 'Codigo' , RTRIM(Des_Color) AS 'DescripciÛn' FROM lb_color where"
    txtCodColor = Trim(txtCodColor)
    txtDesColor = Trim(txtDesColor)

    Select Case opcion
    Case 1: strSQL = strSQL & " Cod_Color   like '%" & Trim(txtCodColor.Text) & "%'  "
    Case 2: strSQL = strSQL & " Des_Color  like '%" & Trim(txtDesColor.Text) & "%' "
    Case 3: strSQL = " SELECT RTRIM(Cod_Color) as 'Codigo' , RTRIM(Des_Color) AS 'DescripciÛn' FROM lb_color"
    End Select
    
    
    txtCodColor = ""
    txtDesColor = ""
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
        .Cargar_Datos
        
        codigo = ""
        Descripcion = ""
        
        iRows = .DGridLista.RowCount
        Set rstAux = .DGridLista.ADORecordset
        If .DGridLista.RowCount > 1 Then
            .Show vbModal
        ElseIf .DGridLista.RowCount = 1 Then
            codigo = .DGridLista.Value(.DGridLista.Columns("CODIGO").Index)
            Descripcion = .DGridLista.Value(.DGridLista.Columns("DESCRIPCIÛN").Index)
        End If
                                
                                
        If codigo <> "" Then
            txtCodColor = RTrim(codigo)
            txtDesColor = RTrim(Descripcion)
            Me.txtCodUniMedida.SetFocus
        End If
            
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub

Private Sub txtKgsProg_GotFocus()
    SelectionText Me.txtKgsProg
End Sub

Private Sub txtKgsProg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtUnidadesProg.SetFocus
    End If
End Sub

Private Sub txtPeso_Bruto_Prog_GotFocus()
    SelectionText Me.txtPeso_Bruto_Prog
End Sub

Private Sub txtKgs_GotFocus()
    'SelectionText Me.txtKgs
End Sub

Private Sub txtPeso_Bruto_Prog_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPeso_Neto_Prog.SetFocus
End If
End Sub
Private Sub txtPeso_Neto_Prog_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtRollosProg.SetFocus
End If
End Sub

Private Sub txtPeso_Neto_Prog_GotFocus()
    SelectionText Me.txtPeso_Neto_Prog
End Sub

Private Sub txtRollosProg_GotFocus()
    SelectionText Me.txtRollosProg
End Sub

Private Sub txtRollosProg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtUbicajeProg.SetFocus
    End If
End Sub

Private Sub txtUbicajeProg_GotFocus()
    SelectionText Me.txtUbicajeProg
End Sub

Private Sub txtUbicajeProg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtKgsProg.SetFocus
    End If
End Sub

Private Sub txtUnidadesProg_GotFocus()
    SelectionText Me.txtUnidadesProg
End Sub

Private Sub GrabaEmbarqueTelas()
On Error GoTo errx
Dim ssql As String

ssql = "TG_Embarque_Telas_man '$',$,$,'$','$','$','$',$,$,$,$,$,$"
  
ssql = VBsprintf(ssql, Saccion, lNum_Embarque, lSec_Embarque, Me.txtCodTela, Me.txtCodComb, Me.txtCodColor, Me.txtCodUniMedida, Me.txtPeso_Bruto_Prog, Me.txtPeso_Neto_Prog, Me.txtRollosProg, Me.txtUbicajeProg, Me.txtKgsProg, Me.txtUnidadesProg)
  

ExecuteCommandSQL cCONNECT, ssql

oParent.BUSCAR
Unload Me

Exit Sub
errx:
    errores err.Number
End Sub


Private Sub txtUnidadesProg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.FunctOKCancel.SetFocus
    End If
End Sub
