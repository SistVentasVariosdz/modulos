VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form frmDetHilCruInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   690
      Left            =   75
      TabIndex        =   14
      Top             =   0
      Width           =   8145
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
         Left            =   3375
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
         Left            =   2925
         TabIndex        =   15
         Top             =   285
         Width           =   285
      End
   End
   Begin VB.Frame Frame2 
      Height          =   810
      Left            =   75
      TabIndex        =   2
      Top             =   5475
      Width           =   8145
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   2070
         TabIndex        =   3
         Top             =   195
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   900
         Custom          =   $"frmDetHilCruInfo.frx":0000
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
      Height          =   4770
      Left            =   75
      TabIndex        =   1
      Top             =   675
      Width           =   8145
      Begin GridEX20.GridEX gexLotes 
         Height          =   3585
         Left            =   60
         TabIndex        =   0
         Top             =   210
         Width           =   7965
         _ExtentX        =   14049
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
         Column(1)       =   "frmDetHilCruInfo.frx":00E7
         Column(2)       =   "frmDetHilCruInfo.frx":01AF
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmDetHilCruInfo.frx":0253
         FormatStyle(2)  =   "frmDetHilCruInfo.frx":038B
         FormatStyle(3)  =   "frmDetHilCruInfo.frx":043B
         FormatStyle(4)  =   "frmDetHilCruInfo.frx":04EF
         FormatStyle(5)  =   "frmDetHilCruInfo.frx":05C7
         FormatStyle(6)  =   "frmDetHilCruInfo.frx":067F
         ImageCount      =   0
         PrinterProperties=   "frmDetHilCruInfo.frx":075F
      End
      Begin VB.Label lblDes_Proveedor 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   3300
         TabIndex        =   13
         Top             =   4290
         Width           =   1680
      End
      Begin VB.Label Label5 
         Caption         =   "Prov."
         Height          =   270
         Left            =   2745
         TabIndex        =   12
         Top             =   4305
         Width           =   525
      End
      Begin VB.Label lblCod_Proveedor 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   3300
         TabIndex        =   11
         Top             =   3945
         Width           =   1680
      End
      Begin VB.Label Label4 
         Caption         =   "C.Prov"
         Height          =   270
         Left            =   2745
         TabIndex        =   10
         Top             =   3975
         Width           =   525
      End
      Begin VB.Label lblStock 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5745
         TabIndex        =   9
         Top             =   4275
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "Stock"
         Height          =   270
         Left            =   5205
         TabIndex        =   8
         Top             =   4305
         Width           =   525
      End
      Begin VB.Label lblCod_Calidad 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   900
         TabIndex        =   7
         Top             =   4290
         Width           =   1680
      End
      Begin VB.Label Label2 
         Caption         =   "Calidad"
         Height          =   270
         Left            =   300
         TabIndex        =   6
         Top             =   4305
         Width           =   525
      End
      Begin VB.Label lblCod_OrdProv 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   900
         TabIndex        =   5
         Top             =   3960
         Width           =   1680
      End
      Begin VB.Label Label1 
         Caption         =   "Lote"
         Height          =   270
         Left            =   300
         TabIndex        =   4
         Top             =   3975
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmDetHilCruInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vCod_TipOrdTra As String, vCod_OrdTra As String, bCancel As Boolean, vCod_Almacen As String, bCancelSec As Boolean

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
        .Columns("COD_HILTEL").Caption = "Cod.Hilado"
        .Columns("DES_HILTEL").Caption = "Hilado"
        .Columns("KGS_PROGR").Caption = "Kgs.Prog."
        .Columns("KGS_CRUDO").Caption = "Kgs.Crudo"
        
        .Columns("NUM_SECUENCIA").Width = 500
        .Columns("COD_HILTEL").Width = 900
        .Columns("DES_HILTEL").Width = 2500
        .Columns("KGS_PROGR").Width = 1000
        .Columns("KGS_CRUDO").Width = 1000
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
            .sQuery = "EXEC TH_SM_BUSCA_CRUDO_SEGUN_HILADO '" & vCod_Almacen & _
            "', '" & gexLotes.Value(gexLotes.Columns("COD_HILTEL").Index) & "'"
            
            .Caption = "Lotes por " & vCod_Almacen & "-" & gexLotes.Value(gexLotes.Columns("COD_HILTEL").Index)
            
            .CARGAR_DATOS
            'Dar Formato al Grid
            .gexLista.Columns("COD_TIPORDTRA").Caption = "Tip.Ord."
            .gexLista.Columns("COD_ORDTRA").Caption = "OT"
            .gexLista.Columns("COD_ORDPROV").Caption = "Lote"
            .gexLista.Columns("COD_CALIDAD").Caption = "Calidad"
            .gexLista.Columns("STOCK").Caption = "Stock"
            .gexLista.Columns("COD_PROVEEDOR").Caption = "Cod.Proveedor"
            .gexLista.Columns("COD_GRUPOTEX").Caption = "Grupo Textil"
            .gexLista.Columns("DES_PROVEEDOR").Caption = "Proveedor"
            
            .gexLista.Columns("COD_TIPORDTRA").Width = 500
            .gexLista.Columns("COD_ORDTRA").Width = 800
            .gexLista.Columns("COD_ORDPROV").Width = 1000
            .gexLista.Columns("COD_CALIDAD").Width = 800
            .gexLista.Columns("STOCK").Width = 1200
            .gexLista.Columns("COD_PROVEEDOR").Width = 2500
            .gexLista.Columns("COD_GRUPOTEX").Width = 1300
            .gexLista.Columns("DES_PROVEEDOR").Width = 2500
            If .gexLista.RowCount > 1 Then .Show vbModal
            bCancelSec = .bCancel
            If .gexLista.RowCount > 0 And Not .bCancel Then
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
End Sub

Private Sub gexLotes_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    LimpiaDet
    bCancelSec = True
End Sub
