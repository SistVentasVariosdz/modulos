VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmOrdCompItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Ordenes de Compra"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   500
      Left            =   9705
      TabIndex        =   2
      Top             =   4485
      Width           =   1100
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   570
      Left            =   60
      TabIndex        =   1
      Top             =   4440
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   1005
      Custom          =   $"frmOrdCompItem.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1200
      ControlHeigth   =   550
      ControlSeparator=   110
   End
   Begin VB.Frame FraLista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      Begin GridEX20.GridEX gexLista 
         Height          =   3930
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   10530
         _ExtentX        =   18574
         _ExtentY        =   6932
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmOrdCompItem.frx":022A
         Column(2)       =   "frmOrdCompItem.frx":02F2
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmOrdCompItem.frx":0396
         FormatStyle(2)  =   "frmOrdCompItem.frx":04CE
         FormatStyle(3)  =   "frmOrdCompItem.frx":057E
         FormatStyle(4)  =   "frmOrdCompItem.frx":0632
         FormatStyle(5)  =   "frmOrdCompItem.frx":070A
         FormatStyle(6)  =   "frmOrdCompItem.frx":07C2
         ImageCount      =   0
         PrinterProperties=   "frmOrdCompItem.frx":08A2
      End
   End
End
Attribute VB_Name = "frmOrdCompItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Strsql As String
Dim Rs_Lista As ADODB.Recordset
Dim sTipo As String

'Definición de Codigo de grupo para ser usado en el ingreso masivo de requerimientos
Public varCod_GrupoTex As String
Public varDes_Grupo As String

'Definicion de variables que seran pasadas por nuestro master
Public varSer_OrdComp As String, varCod_OrdComp As String, varSec_OrdComp As String
Public varTip_Presentacion As String, varCod_ClaOrdComp As String, varCod_Proveedor As String
Public varCod_Descuento As String
Public varPorc_IGV As Double
Public varCod_TipRequ As Integer
Public varCod_StaOrdComp As String

Dim Tipo_Consulta As String
Sub CARGA_GRID()
    
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = cConnect
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    Strsql = "SELECT Tip_Item FROM lg_claordcomp WHERE Cod_ClaOrdComp = '" & varCod_ClaOrdComp & "'"
    Tipo_Consulta = DevuelveCampo(Strsql, cConnect)
    
    If Tipo_Consulta = "" Then
        Tipo_Consulta = "I"
    End If
    
    'Esta cadena es para devolver el Codigo de Cliente
    Strsql = "EXEC UP_SEL_ORDCOMPITEM '" & Tipo_Consulta & "','" & varSer_OrdComp & "','" & varCod_OrdComp & "'"
    
    Rs_Lista.Open Strsql
    'Set DGridLista.DataSource = Rs_Lista
    Set gexLista.ADORecordset = Rs_Lista 'CargarRecordSetDesconectado(Strsql, cConnect)
    gexLista.Refresh

    If gexLista.RowCount > 0 Then
        gexLista.Enabled = True
        HabilitaMant Me.FunctButt1, "ADICIONAR/MODIFICAR/ELIMINAR/REQUERIMIENTOS/CERRAR/ENTREGAS/IMPRIMIR"
        'Call CARGA_DATOS
    Else
        gexLista.Enabled = False
        HabilitaMant Me.FunctButt1, "ADICIONAR"
        'Call LIMPIAR_DATOS
    End If
    
    gexLista.Columns("Item").Width = 2500
    gexLista.Columns("Combinacion").Width = 0
    gexLista.Columns("Color").Width = 1800
    'gexLista.Columns("MEDIDA").Width = 700
    gexLista.Columns("Destino").Width = 0
    gexLista.Columns("Estilo_Cli").Width = 0
    gexLista.Columns("Cod_UniMed").Width = 800
    gexLista.Columns("Pre_Unitario").Width = 900
    gexLista.Columns("Can_Requerida").Width = 1000
    gexLista.Columns("Can_Comprada").Width = 1000
    gexLista.Columns("Can_Recibida").Width = 1000
    gexLista.Columns("Fac_EquiProv").Width = 1000
    gexLista.Columns("Cod_ItemProv").Width = 1000
    gexLista.Columns("Cod_Prov").Width = 1000
    gexLista.Columns("Observaciones").Width = 2000
    
    gexLista.Columns("Estilo_Cli").Caption = "Est. Cliente"
    gexLista.Columns("Cod_UniMed").Caption = "Uni. Med."
    gexLista.Columns("Pre_Unitario").Caption = "P. Unit."
    gexLista.Columns("Can_Requerida").Caption = "C. Requerida"
    gexLista.Columns("Can_Comprada").Caption = "C. Comprada"
    gexLista.Columns("Can_Recibida").Caption = "C. Recibida"
    gexLista.Columns("Fac_EquiProv").Caption = "F. Equivalencia"
    gexLista.Columns("Cod_ItemProv").Caption = "Item Prov."
    gexLista.Columns("Cod_Prov").Caption = "Cod Prov"
        
    gexLista.Columns("Ser_OrdComp").Width = 0
    gexLista.Columns("cod_OrdComp").Width = 0
    gexLista.Columns("Sec_OrdComp").Width = 0
    gexLista.Columns("codigo").Width = 0
    gexLista.Columns("Descripcion").Width = 0
    gexLista.Columns("cod_comb").Width = 0
    gexLista.Columns("Des_comb").Width = 0
    gexLista.Columns("cod_Color").Width = 0
    gexLista.Columns("Des_Color").Width = 0
    gexLista.Columns("Cod_Talla").Width = 0
    gexLista.Columns("cod_Destino").Width = 0
    gexLista.Columns("Des_Destino").Width = 0
    gexLista.Columns("cod_EstCli").Width = 0
    gexLista.Columns("Des_EstCli").Width = 0
    gexLista.Columns("cod_descuento").Width = 0
    gexLista.Columns("porc_igv").Width = 0
    gexLista.Columns("fec_entrega_inicio").Width = 0
    gexLista.Columns("fec_entrega_fin").Width = 0
    
    gexLista.FrozenColumns = 14
End Sub

Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    If sTipo <> "D" Then
    Else
        'Aqui se valida que no tenga registros dependientes
        'Strsql = "SELECT COUNT(*) FROM LG_ORDCOMPITEMREQ WHERE Ser_OrdComp = '" & varSer_OrdComp & "' AND Cod_OrdComp = '" & varCod_OrdComp & "' AND Sec_OrdComp = '" & Rs_Lista("Sec_OrdComp").Value & "'"
        Strsql = "SELECT COUNT(*) FROM LG_ORDCOMPITEMREQ WHERE Ser_OrdComp = '" & varSer_OrdComp & "' AND Cod_OrdComp = '" & varCod_OrdComp & "' AND Sec_OrdComp = '" & gexLista.Value(gexLista.Columns("Sec_OrdComp").Index) & "'"
        
        If DevuelveCampo(Strsql, cConnect) > 0 Then
            MsgBox "El registro no puede ser eliminado por que posee registros relacionados. Sirvase verificar", vbInformation, "Ordenes de Compra"
            VALIDA_DATOS = False
            Exit Function
        End If
    End If
End Function

Sub ELIMINAR_DATOS()
 Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
   
    Con.ConnectionString = cConnect
    Con.Open
    Con.BeginTrans
       
        'Strsql = "EXEC UP_MAN_ORDCOMPITEMN '" & _
        sTipo & "','" & _
        varSer_OrdComp & "','" & _
        varCod_OrdComp & "','" & _
        Rs_Lista("Sec_OrdComp").Value & "','" & _
        "" & "','" & _
        "" & "','" & _
        "" & "','" & _
        "" & "','" & _
        "" & "','" & _
        "" & "','" & _
        "" & "'," & _
        "0" & ",'" & _
        "" & "','" & _
        "" & "','" & _
        "0" & "','" & _
        "0" & "','" & _
        "" & "','','',''"
        Strsql = "EXEC UP_MAN_ORDCOMPITEMN '" & _
        sTipo & "','" & _
        varSer_OrdComp & "','" & _
        varCod_OrdComp & "','" & _
        gexLista.Value(gexLista.Columns("Sec_OrdComp").Index) & "','" & _
        "" & "','" & _
        "" & "','" & _
        "" & "','" & _
        "" & "','" & _
        "" & "','" & _
        "" & "','" & _
        "" & "'," & _
        "0" & ",'" & _
        "" & "','" & _
        "" & "','" & _
        "0" & "','" & _
        "0" & "','" & _
        "" & "','','',''"
        
        
        Con.Execute Strsql
    
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMESSAGE_INF_DATA_DELETE
    Informa "", amensaje
    
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"

End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Call FormateaGrid(DGridLista)
    Me.FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    'Call CARGA_GRID
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim eliminar As Integer
    Select Case ActionName
        Case "ADICIONAR":
        
                        If varCod_StaOrdComp <> "P" Then
                            MsgBox "El estado del registro no permite modificación alguna. Sirvase verificar", vbInformation, "Ordenes de Compra"
                            Exit Sub
                        End If
        
                        Strsql = "SELECT Flg_Requerimiento FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Me.varCod_ClaOrdComp & "'"
                        If DevuelveCampo(Strsql, cConnect) = "N" Then
                            Strsql = "SELECT Tip_Item FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Me.varCod_ClaOrdComp & "'"
                            
                            frmOrdCompItemAddN.varTip_Presentacion = DevuelveCampo(Strsql, cConnect)
                            frmOrdCompItemAddN.varSer_OrdComp = Me.varSer_OrdComp
                            frmOrdCompItemAddN.varCod_OrdComp = Me.varCod_OrdComp
                            frmOrdCompItemAddN.varCod_ClaOrdComp = Me.varCod_ClaOrdComp
                            frmOrdCompItemAddN.varCod_StaOrdComp = Me.varCod_StaOrdComp
                            frmOrdCompItemAddN.varPorc_IGV = Me.varPorc_IGV
                            frmOrdCompItemAddN.varCod_Descuento = Me.varCod_Descuento
                            frmOrdCompItemAddN.varCod_Proveedor = Me.varCod_Proveedor
                            Load frmOrdCompItemAddN
                            frmOrdCompItemAddN.CARGA_GRID
                            frmOrdCompItemAddN.Show 1
                        Else
                            
                            If varCod_TipRequ = 1 Then
                                Load FrmReqCompra
                                
                                FrmReqCompra.varCod_TipRequ = Me.varCod_TipRequ
                                FrmReqCompra.varTip_Presentacion = DevuelveCampo(Strsql, cConnect)
                                FrmReqCompra.varSer_OrdComp = Me.varSer_OrdComp
                                FrmReqCompra.varCod_OrdComp = Me.varCod_OrdComp
                                FrmReqCompra.varCod_ClaOrdComp = Me.varCod_ClaOrdComp
                                FrmReqCompra.varPorc_IGV = Me.varPorc_IGV
                                FrmReqCompra.varCod_Proveedor = Me.varCod_Proveedor
                                FrmReqCompra.varCod_GrupoLog = Me.varCod_GrupoTex
                                FrmReqCompra.TxtDes_Grupo = varDes_Grupo
                                'Call FrmReqCompra.CARGA_LISTA(Me.varCod_TipRequ)
                            
                                FrmReqCompra.Show 1
                            Else
                                Load frmOrdCompItemAddS
                                frmOrdCompItemAddS.varCod_TipRequ = Me.varCod_TipRequ
                                
                                frmOrdCompItemAddS.varTip_Presentacion = DevuelveCampo(Strsql, cConnect)
                                frmOrdCompItemAddS.varSer_OrdComp = Me.varSer_OrdComp
                                frmOrdCompItemAddS.varCod_OrdComp = Me.varCod_OrdComp
                                If Rs_Lista.RecordCount > 0 Then
                                    'frmOrdCompItemAddS.varSec_OrdComp = Rs_Lista("Sec_OrdComp").Value
                                    frmOrdCompItemAddS.varSec_OrdComp = gexLista.Value(gexLista.Columns("Sec_OrdComp").Index)
                                End If
                                frmOrdCompItemAddS.varCod_ClaOrdComp = Me.varCod_ClaOrdComp
                                frmOrdCompItemAddS.varPorc_IGV = Me.varPorc_IGV
                                frmOrdCompItemAddS.varCod_Proveedor = Me.varCod_Proveedor
                                frmOrdCompItemAddS.varCod_GrupoTex = Me.varCod_GrupoTex
                                frmOrdCompItemAddS.TxtDes_Grupo = varDes_Grupo
                                
                                Call frmOrdCompItemAddS.MUESTRA_GRID(Me.varCod_TipRequ)
                                frmOrdCompItemAddS.Show 1
                            End If
                        End If
                        Call CARGA_GRID
                        
        Case "MODIFICAR":
                        'If varCod_StaOrdComp <> "P" Then
                        '    MsgBox "El estado del registro no permite modificación alguna. Sirvase verificar", vbInformation, "Ordenes de Compra"
                        '    Exit Sub
                        'End If

'                        Strsql = "SELECT Flg_Requerimiento FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Me.varCod_ClaOrdComp & "'"
'                        If DevuelveCampo(Strsql, cCONNECT) = "N" Then
                            Strsql = "SELECT Tip_Item FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Me.varCod_ClaOrdComp & "'"
                            Load frmOrdCompItemAddN
                            frmOrdCompItemAddN.varCod_StaOrdComp = Me.varCod_StaOrdComp
                            frmOrdCompItemAddN.varTip_Presentacion = DevuelveCampo(Strsql, cConnect)
                            frmOrdCompItemAddN.varSer_OrdComp = Me.varSer_OrdComp
                            frmOrdCompItemAddN.varCod_OrdComp = Me.varCod_OrdComp
                            frmOrdCompItemAddN.varCod_ClaOrdComp = Me.varCod_ClaOrdComp
                            frmOrdCompItemAddN.varPorc_IGV = Me.varPorc_IGV
                            frmOrdCompItemAddN.varCod_Descuento = Me.varCod_Descuento
                            frmOrdCompItemAddN.varCod_Proveedor = Me.varCod_Proveedor
                            frmOrdCompItemAddN.CARGA_GRID
                            frmOrdCompItemAddN.Show 1
'                        Else
'
'                            If varCod_TipRequ = 1 Then
'                                FrmReqCompra.Show 1
'                            Else
'                                Load frmOrdCompItemAddS
'                                frmOrdCompItemAddS.varCod_TipRequ = Me.varCod_TipRequ
'
'                                frmOrdCompItemAddS.varTip_Presentacion = DevuelveCampo(Strsql, cCONNECT)
'                                frmOrdCompItemAddS.varSer_OrdComp = Me.varSer_OrdComp
'                                frmOrdCompItemAddS.varCod_OrdComp = Me.varCod_OrdComp
'                                frmOrdCompItemAddS.varCod_ClaOrdComp = Me.varCod_ClaOrdComp
'                                frmOrdCompItemAddS.varPorc_IGV = Me.varPorc_IGV
'                                frmOrdCompItemAddS.varCod_Proveedor = Me.varCod_Proveedor
'
'
'                                Call frmOrdCompItemAddS.CARGA_LISTA(Me.varCod_TipRequ)
'                                frmOrdCompItemAddS.Show 1
'                            End If
'                        End If
                        Call CARGA_GRID
                        
                        
        Case "ELIMINAR":
                        
                        If varCod_StaOrdComp <> "P" Then
                            MsgBox "El estado del registro no permite modificación alguna. Sirvase verificar", vbInformation, "Ordenes de Compra"
                            Exit Sub
                        End If
                        eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Ordenes de Compra")
                        If eliminar = vbYes Then
                        sTipo = "D"
                            If VALIDA_DATOS Then
                                Call ELIMINAR_DATOS
                                Call CARGA_GRID
                                sTipo = ""
                            End If
                        End If

        Case "CERRAR":
                        'MsgBox ("D")
                        
        Case "REQUERIMIENTOS":
                        Load frmReqDet
                        frmReqDet.varSer_OrdComp = Me.varSer_OrdComp
                        frmReqDet.varCod_OrdComp = Me.varCod_OrdComp
                        'frmReqDet.varSec_OrdComp = Rs_Lista("Sec_OrdComp").Value
                        frmReqDet.varSec_OrdComp = gexLista.Value(gexLista.Columns("Sec_OrdComp").Index)
                        frmReqDet.varCod_ClaOrdComp = Me.varCod_ClaOrdComp
                        frmReqDet.varCod_StaOrdComp = Me.varCod_StaOrdComp
                        frmReqDet.CARGA_GRID
                        frmReqDet.Show 1
                        
                        Call CARGA_GRID
                        
        Case "ENTREGAS"
                        Load frmVerEntregas
                        'frmVerEntregas.Caption = "Entregas de la Orden de Compra: " & Me.varSer_OrdComp & " - " & Me.varCod_OrdComp & "  Secuencia : " & Rs_Lista("Sec_OrdComp").Value
                        frmVerEntregas.Caption = "Entregas de la Orden de Compra: " & Me.varSer_OrdComp & " - " & Me.varCod_OrdComp & "  Secuencia : " & gexLista.Value(gexLista.Columns("Sec_OrdComp").Index)
                        frmVerEntregas.varSer_OrdComp = Me.varSer_OrdComp
                        frmVerEntregas.varCod_OrdComp = Me.varCod_OrdComp
                        'frmVerEntregas.varSec_OrdComp = Rs_Lista("Sec_OrdComp").Value
                        frmVerEntregas.varSec_OrdComp = gexLista.Value(gexLista.Columns("Sec_OrdComp").Index)
                        frmVerEntregas.CARGA_GRID
                        frmVerEntregas.Show 1
        Case "IMPRIMIR"
            Call REPORTE
    End Select
End Sub


Public Sub REPORTE()
On Error GoTo ErrorImpresion
    Dim oo As Object
    Set oo = CreateObject("excel.application")
    
    oo.Workbooks.Open vRuta & "\OrdCompDetalle.xlt"
    oo.Visible = True
    oo.Run "REPORTE", Tipo_Consulta, Me.varSer_OrdComp, Me.varCod_OrdComp, cConnect
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte de Seguimiento de O/C" & Err.Description, vbCritical, "Impresion"
End Sub

