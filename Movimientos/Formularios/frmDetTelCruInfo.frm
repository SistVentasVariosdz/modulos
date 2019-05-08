VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmDetTelCruInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   690
      Left            =   75
      TabIndex        =   14
      Top             =   0
      Width           =   9915
      Begin VB.Label lblOT 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   16
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label6 
         Caption         =   "OT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3390
         TabIndex        =   15
         Top             =   285
         Width           =   285
      End
   End
   Begin VB.Frame Frame2 
      Height          =   810
      Left            =   75
      TabIndex        =   2
      Top             =   6120
      Width           =   9915
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   2820
         TabIndex        =   3
         Top             =   195
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   900
         Custom          =   $"frmDetTelCruInfo.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5310
      Left            =   90
      TabIndex        =   1
      Top             =   810
      Width           =   9870
      Begin GridEX20.GridEX gexLotes 
         Height          =   3585
         Left            =   60
         TabIndex        =   0
         Top             =   210
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   6324
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         DataMode        =   1
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmDetTelCruInfo.frx":00E7
         Column(2)       =   "frmDetTelCruInfo.frx":01AF
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmDetTelCruInfo.frx":0253
         FormatStyle(2)  =   "frmDetTelCruInfo.frx":038B
         FormatStyle(3)  =   "frmDetTelCruInfo.frx":043B
         FormatStyle(4)  =   "frmDetTelCruInfo.frx":04EF
         FormatStyle(5)  =   "frmDetTelCruInfo.frx":05C7
         FormatStyle(6)  =   "frmDetTelCruInfo.frx":067F
         ImageCount      =   0
         PrinterProperties=   "frmDetTelCruInfo.frx":075F
      End
      Begin VB.Label lblDes_Tela 
         Height          =   285
         Left            =   6930
         TabIndex        =   20
         Top             =   4950
         Width           =   2310
      End
      Begin VB.Label lblCod_Comb 
         Height          =   240
         Left            =   2070
         TabIndex        =   19
         Top             =   4905
         Width           =   1590
      End
      Begin VB.Label lblCod_Tela 
         Height          =   285
         Left            =   4455
         TabIndex        =   18
         Top             =   4815
         Width           =   1635
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1680
      End
      Begin VB.Label lblDes_Proveedor 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   4410
         TabIndex        =   13
         Top             =   4290
         Width           =   1680
      End
      Begin VB.Label Label5 
         Caption         =   "Prov."
         Height          =   270
         Left            =   3870
         TabIndex        =   12
         Top             =   4305
         Width           =   525
      End
      Begin VB.Label lblCod_Proveedor 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   4425
         TabIndex        =   11
         Top             =   3945
         Width           =   1680
      End
      Begin VB.Label Label4 
         Caption         =   "C.Prov"
         Height          =   270
         Left            =   3870
         TabIndex        =   10
         Top             =   3975
         Width           =   525
      End
      Begin VB.Label lblStock 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   6870
         TabIndex        =   9
         Top             =   4275
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "Stock"
         Height          =   270
         Left            =   6330
         TabIndex        =   8
         Top             =   4305
         Width           =   525
      End
      Begin VB.Label lblCod_Calidad 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   2025
         TabIndex        =   7
         Top             =   4290
         Width           =   1680
      End
      Begin VB.Label Label2 
         Caption         =   "Calidad"
         Height          =   270
         Left            =   1425
         TabIndex        =   6
         Top             =   4305
         Width           =   525
      End
      Begin VB.Label lblCod_OrdProv 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   2025
         TabIndex        =   5
         Top             =   3960
         Width           =   1680
      End
      Begin VB.Label Label1 
         Caption         =   "Lote"
         Height          =   270
         Left            =   1425
         TabIndex        =   4
         Top             =   3975
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmDetTelCruInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vCod_TipOrdTra As String, vCod_OrdTra As String, vNum_Secuencia As Integer, bCancel As Boolean, vCod_Almacen As String, bCancelSec As Boolean

Public Sub SM_AYUDA_ITEMS_DE_PARTIDA()
Dim strSQL As String
    bCancel = False
    LimpiaDet
    lblOT = vCod_TipOrdTra & " - " & vCod_OrdTra
    strSQL = "EXEC SM_AYUDA_ITEMS_DE_PARTIDA '" & vCod_TipOrdTra & "', '" & _
             vCod_OrdTra & "'"
    Set gexLotes.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    With gexLotes
        .Columns("NUM_SECUENCIA").Caption = "Sec."
        .Columns("COD_TELA").Caption = "Cod.Tela"
        .Columns("DES_TELA").Caption = "Tela"
        .Columns("COD_COMB").Caption = "Cod.Comb"
        .Columns("DES_COMB").Caption = "Combinación"
        .Columns("COD_TALLA").Caption = "Cod.Talla"
        .Columns("MEDIDA").Caption = "Medida"
        .Columns("KGS_PROGR").Caption = "Kgs.Prog."
        .Columns("KGS_CRUDO").Caption = "Kgs.Crudo"
        .Columns("UNI_CRUDO").Caption = "Und.Crudo"
        .Columns("NUM_BULTOS_ENVIADOS").Caption = "Bultos Env."
        
        .Columns("NUM_SECUENCIA").Width = 465
        .Columns("COD_TELA").Width = 945
        .Columns("DES_TELA").Width = 2505
        .Columns("COD_COMB").Visible = False
        .Columns("DES_COMB").Width = 1005
        .Columns("COD_TALLA").Width = 495
        .Columns("MEDIDA").Width = 660
        .Columns("KGS_PROGR").Width = 825
        .Columns("KGS_CRUDO").Width = 870
        .Columns("UNI_CRUDO").Width = 900
        .Columns("NUM_BULTOS_ENVIADOS").Width = 945
    End With
    If gexLotes.RowCount = 1 Then FunctButt1_ActionClick 0, 0, "LOTE"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim strSQL As String
    Select Case ActionName
    Case "ACEPTAR"
        Me.Hide
    Case "LOTE"
        If gexLotes.RowCount = 0 Then Exit Sub
        With frmBusqGeneral3
            .sQuery = "EXEC SM_BUSCA_CRUDO_SEGUN_TELA_COMB_TALLA '" & _
            vCod_TipOrdTra & "', '" & vCod_OrdTra & "', " & _
            gexLotes.Value(gexLotes.Columns("Num_Secuencia").Index) & ", '" & _
            vCod_Almacen & "', '" & gexLotes.Value(gexLotes.Columns("COD_TELA").Index) & _
            "', '" & gexLotes.Value(gexLotes.Columns("COD_COMB").Index) & "', '" & _
            gexLotes.Value(gexLotes.Columns("COD_TALLA").Index) & "'"
            
            .Caption = "Lotes por " & vCod_Almacen & "-" & _
                       gexLotes.Value(gexLotes.Columns("COD_TELA").Index) & "-" & _
                       gexLotes.Value(gexLotes.Columns("COD_COMB").Index) & "-" & _
                       gexLotes.Value(gexLotes.Columns("COD_TALLA").Index)
            
            .Cargar_Datos
            'Dar Formato al Grid
            .gexLista.Columns("COD_TIPORDTRA").Caption = "Tip.Ord."
            .gexLista.Columns("COD_ORDTRA").Caption = "OT"
            .gexLista.Columns("COD_ORDPROV").Caption = "Lote"
            .gexLista.Columns("COD_CALIDAD").Caption = "Calidad"
            .gexLista.Columns("STOCK").Caption = "Stock"
            .gexLista.Columns("COD_PROVEEDOR").Caption = "Cod.Proveedor"
            .gexLista.Columns("COD_GRUPOTEX").Caption = "Grupo Textil"
            .gexLista.Columns("DES_PROVEEDOR").Caption = "Proveedor"
            
            .gexLista.Columns("COD_TIPORDTRA").Width = 495
            .gexLista.Columns("COD_ORDTRA").Width = 435
            .gexLista.Columns("COD_ORDPROV").Width = 1590
            .gexLista.Columns("COD_CALIDAD").Width = 270
            .gexLista.Columns("STOCK").Width = 555
            .gexLista.Columns("COD_PROVEEDOR").Width = 1500
            .gexLista.Columns("COD_GRUPOTEX").Width = 915
            .gexLista.Columns("DES_PROVEEDOR").Width = 1455
            If .gexLista.RowCount > 1 Then .Show vbModal
            bCancelSec = .bCancel
            If .gexLista.RowCount > 0 And Not .bCancel Then
                lblCod_Tela = .gexLista.Value(.gexLista.Columns("COD_tela").Index)
                lblDes_Tela = .gexLista.Value(.gexLista.Columns("des_tela").Index)
                lblCod_Comb = .gexLista.Value(.gexLista.Columns("COD_comb").Index)
                lblCod_OrdProv = .gexLista.Value(.gexLista.Columns("COD_ORDPROV").Index)
                lblCod_Calidad = .gexLista.Value(.gexLista.Columns("COD_CALIDAD").Index)
                lblStock = .gexLista.Value(.gexLista.Columns("STOCK").Index)
                lblCod_Proveedor = .gexLista.Value(.gexLista.Columns("COD_PROVEEDOR").Index)
                lblDes_Proveedor = .gexLista.Value(.gexLista.Columns("DES_PROVEEDOR").Index)
            End If
        End With
        Unload frmBusqGeneral3
    Case "CANCELAR"
        Me.Hide
        bCancel = True
    End Select
End Sub

Private Sub gexLotes_DblClick()
    FunctButt1_ActionClick 0, 0, "LOTE"
End Sub

Private Sub gexLotes_KeyDown(KeyCode As Integer, Shift As Integer)
'Al presionar Enter la fila baja y no se selecciona el regsitro deseado
'    Select Case KeyCode
'    Case vbKeyReturn
'        FunctButt1_ActionClick 0, 0, "LOTE"
'    Case vbKeyEscape
'        FunctButt1_ActionClick 0, 0, "CANCELAR"
'    End Select
End Sub

Private Sub LimpiaDet()
    lblCod_OrdProv = "": lblCod_Calidad = ""
    lblStock = "": lblCod_Proveedor = ""
    lblDes_Proveedor = ""
    lblCod_Tela = ""
    lblCod_Comb = ""
    lblDes_Tela = ""
End Sub


Private Sub gexLotes_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    LimpiaDet
    bCancelSec = True
End Sub
