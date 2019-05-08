VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} RptOCompra 
   ClientHeight    =   10020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   OleObjectBlob   =   "RptOCompra.dsx":0000
End
Attribute VB_Name = "RptOCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public varSer_OrdComp As String
Public varCod_OrdComp As String
Public varFormulado As String
Public varGrupo As String
Public varTipItem As String
Dim Strsql As String


Sub Carga_Reporte()
Dim Rs_Lista As New ADODB.Recordset

    Rs_Lista.ActiveConnection = cConnect
    Rs_Lista.CursorLocation = adUseClient
    
    Strsql = "EXEC UP_SEL_RPTOCOMPRACAB '" & varSer_OrdComp & "','" & varCod_OrdComp & "'"
    Set Rs_Lista = CargarRecordSetDesconectado(Strsql, cConnect)
    Me.Database.SetDataSource Rs_Lista, , 1
    
    Strsql = "EXEC UP_SEL_ORDCOMPITEMREP_SIMPLE '" & varTipItem & "','" & varSer_OrdComp & "','" & varCod_OrdComp & "'"
    Set Rs_Lista = CargarRecordSetDesconectado(Strsql, cConnect)
    'm_Report.Database.SetDataSource Rs_Lista, , 2
    Me.Database.SetDataSource Rs_Lista, , 2
End Sub


