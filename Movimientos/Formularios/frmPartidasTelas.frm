VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBusqPartidasTelas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda Partidas Telas"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   510
      Left            =   2250
      TabIndex        =   1
      Top             =   3195
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   510
      Left            =   3840
      TabIndex        =   2
      Top             =   3195
      Width           =   1395
   End
   Begin GridEX20.GridEX gexLista 
      Height          =   3000
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5292
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmPartidasTelas.frx":0000
      Column(2)       =   "frmPartidasTelas.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmPartidasTelas.frx":016C
      FormatStyle(2)  =   "frmPartidasTelas.frx":02A4
      FormatStyle(3)  =   "frmPartidasTelas.frx":0354
      FormatStyle(4)  =   "frmPartidasTelas.frx":0408
      FormatStyle(5)  =   "frmPartidasTelas.frx":04E0
      FormatStyle(6)  =   "frmPartidasTelas.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmPartidasTelas.frx":0678
   End
End
Attribute VB_Name = "frmBusqPartidasTelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Strsql As String
Public varCod_TipOrdTra As String
Public varCod_Ordtra As String

Public oParent As Object
Public varBusqueda As String

Sub CARGA_GRID()
    
    'Esta cadena es para devolver el Codigo de Cliente
    Strsql = "EXEC UP_SEL_TELAS_TX_ORDTRA_ITEMS '" & Me.varCod_TipOrdTra & "','" & Me.varCod_Ordtra & "'"
    
    Set gexLista.ADORecordset = CargarRecordSetDesconectado(Strsql, cConnect)
    
    SetGeneralGridEX gexLista, 0, 1
    
    Call Me.gexLista.Find(4, jgexEqual, varBusqueda)
    'If gexLista.RowCount > 0 Then
    '    gexLista.MoveFirst
    'End If
    
    Call CONFIGURAR_GRID
    
End Sub

Private Sub cmdAceptar_Click()
    If gexLista.RowCount > 0 Then
        With oParent
            .Codigo = gexLista.Value(gexLista.Columns("Cod_Tela").Index)
            .Descripcion = ""
            
            .TxtItem.Text = gexLista.Value(gexLista.Columns("Cod_Tela").Index)
            .TxtDesitem.Text = gexLista.Value(gexLista.Columns("Des_Tela").Index)
            
            .Label3.Caption = gexLista.Value(gexLista.Columns("Des_Comb").Index)
            .Label4.Caption = gexLista.Value(gexLista.Columns("Des_Color").Index)
            .Label5.Caption = gexLista.Value(gexLista.Columns("Cod_Talla").Index)
            .Cod_Talla = gexLista.Value(gexLista.Columns("Cod_Talla").Index)
            .Cod_color = gexLista.Value(gexLista.Columns("Cod_Color").Index)
            .Cod_Comb = FixNulos(gexLista.Value(gexLista.Columns("Cod_Comb").Index), vbString)
        End With
    End If
    Unload Me
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub gexLista_DblClick()
    Call cmdAceptar_Click
End Sub

Private Sub gexLista_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call cmdAceptar_Click
    End If
End Sub

Private Sub gexLista_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then
    '    Call cmdAceptar_Click
    'End If
End Sub

Public Sub CONFIGURAR_GRID()
    Me.gexLista.Columns("Cod_TipOrdTra").Visible = False
    Me.gexLista.Columns("Cod_Ordtra").Visible = False
    Me.gexLista.Columns("Num_Secuencia").Visible = False
    Me.gexLista.Columns("Cod_Tela").Visible = False
    Me.gexLista.Columns("Des_Tela").Visible = False
    Me.gexLista.Columns("Cod_Comb").Visible = False
    Me.gexLista.Columns("Des_Comb").Visible = False
    Me.gexLista.Columns("Cod_Color").Visible = False
    Me.gexLista.Columns("Des_Color").Visible = False

    Me.gexLista.Columns("TELA").Caption = "Tela"
    Me.gexLista.Columns("TELA").Width = 2500
    Me.gexLista.Columns("COMBINACION").Caption = "Combinación"
    Me.gexLista.Columns("COMBINACION").Width = 2500
    Me.gexLista.Columns("COLOR").Caption = "Color"
    Me.gexLista.Columns("COLOR").Width = 2000
    Me.gexLista.Columns("Cod_Talla").Caption = "Talla"
    Me.gexLista.Columns("Cod_Talla").Width = 1000

End Sub



