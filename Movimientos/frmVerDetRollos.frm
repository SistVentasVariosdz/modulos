VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmVerDetRollos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle De Rollos"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   525
      Left            =   0
      TabIndex        =   1
      Top             =   3240
      Width           =   1185
   End
   Begin GridEX20.GridEX gexMovDet 
      Height          =   3285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   5794
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GridLineStyle   =   2
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      HeaderFontName  =   "Verdana"
      HeaderFontBold  =   -1  'True
      HeaderFontSize  =   6.75
      HeaderFontWeight=   700
      FontName        =   "Tahoma"
      ColumnHeaderHeight=   270
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmVerDetRollos.frx":0000
      Column(2)       =   "frmVerDetRollos.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmVerDetRollos.frx":016C
      FormatStyle(2)  =   "frmVerDetRollos.frx":0294
      FormatStyle(3)  =   "frmVerDetRollos.frx":0344
      FormatStyle(4)  =   "frmVerDetRollos.frx":03F8
      FormatStyle(5)  =   "frmVerDetRollos.frx":04D0
      FormatStyle(6)  =   "frmVerDetRollos.frx":0588
      ImageCount      =   0
      PrinterProperties=   "frmVerDetRollos.frx":0668
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   4800
      TabIndex        =   2
      Top             =   3240
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   900
      Custom          =   $"frmVerDetRollos.frx":0840
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   25
   End
End
Attribute VB_Name = "frmVerDetRollos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Scod_ordtra As String
Public sCod_Almacen As String, sNum_MovStk As String, sCod_TipMov As String, sNum_Secuencia As String, sAlmacen As String
Dim strSQL As String
Public sCOD_TIPMOV_ROLLOS_TEJEDURIA As String

Dim sFlg_Devolucion_Rollos_Tejeduria As String

Private Sub CmdImprimir_Click()
    On Error GoTo hand

    Dim oo As Object, sRuta As String
    Dim rs As New Recordset
    strSQL = "EXEC TJ_SM_MUESTRA_MOV_TELA_CRUDA_ROLLOS_REPORTE '" & sCod_Almacen & "', '" & sNum_MovStk & "','" & sNum_Secuencia & "'"

    Set rs = CargarRecordSetDesconectado(strSQL, cConnect)
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\rptDetalleRollos.xlt"
    oo.Visible = True
    
    oo.Run "Reporte", rs
Exit Sub
hand:
ErrorHandler err, Me.Caption
End Sub


Public Sub BUSCAR()
    On Error GoTo Fin
    Dim sTit As String
    sTit = "Mostrar Movimentos Stock Rollos"
    
    strSQL = "EXEC Tj_SM_MUESTRA_MOV_TELA_CRUDA_ROLLOS '" & sCod_Almacen & "', '" & sNum_MovStk & "','" & sNum_Secuencia & "'"
    Set gexMovDet.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
    Dim C As Integer
    
    With gexMovDet
        For C = 1 To .Columns.Count
            With .Columns(C)
                .Caption = UCase(Trim(.Caption))
                .HeaderAlignment = jgexAlignCenter
            End With
        Next C
        .Columns("Cod_Almacen").Visible = False
        .Columns("Num_MovStk").Visible = False
        .Columns("Num_Secuencia_OT").Visible = False
        .Columns("Cod_TipMov").Visible = False
        
        .Columns("Cod_Calidad").Visible = False
 '       .Columns("Cod_Calidad_Auditoria").Visible = False
       ' .Columns("GRUPO_OT").Visible = False
        '.Columns("GRUPO_MAQUINA").Visible = False
        .Columns("SEC_MAQUINA").Visible = False
        .Columns("PREFIJO_MAQUINA").Visible = False
       ' .Columns("FLG_CALIDAD_DIFERENTE").Visible = False
        
        .Columns("Num_Secuencia").Caption = "SEC.MOV"
        .Columns("Num_Secuencia").Visible = False
        
        .Columns("Cod_OrdTra").Caption = "O.T."
        .Columns("Cod_OrdTra").Width = 600
        .Columns("Cod_OrdTra").TextAlignment = jgexAlignCenter
        .Columns("Cod_OrdTra").Visible = False
        
        .Columns("Num_Rollo").Visible = False
        
        .Columns("SEC_MAQUINA").Caption = "SEC.MAQ"
        .Columns("SEC_MAQUINA").TextAlignment = jgexAlignCenter
        .Columns("SEC_MAQUINA").Width = 800
        
        .Columns("Prefijo_Maquina").Caption = "PREF.MAQ"
        .Columns("Prefijo_Maquina").TextAlignment = jgexAlignCenter
        .Columns("Prefijo_Maquina").Width = 900
        
        .Columns("Codigo_Rollo").Caption = "COD.ROLLO"
        .Columns("Codigo_Rollo").TextAlignment = jgexAlignCenter
        .Columns("Codigo_Rollo").Width = 1000
        
        .Columns("Kgs_Rollo").Caption = "Kgs.ROLLO"
        .Columns("Kgs_Rollo").TextAlignment = jgexAlignRight
        .Columns("Kgs_Rollo").Width = 1000
        
        .Columns("Uni_Rollos").Caption = "Uni.ROLLO"
        .Columns("Uni_Rollos").TextAlignment = jgexAlignRight
        .Columns("Uni_Rollos").Width = 1000
        
     '   .Columns("CALIDAD_REG_MOV").Caption = "CALIDAD DEL MOVIMIENTO"
     '   .Columns("CALIDAD_REG_MOV").TextAlignment = jgexAlignLeft
     '   .Columns("CALIDAD_REG_MOV").Width = 2200
        
    '    .Columns("CALIDAD_AUDITORIA").Caption = "CALIDAD DE AUDITORIA"
    '    .Columns("CALIDAD_AUDITORIA").TextAlignment = jgexAlignLeft
    '    .Columns("CALIDAD_AUDITORIA").Width = 2000
        
   '     .Columns("AUDITOR").Caption = "AUDITOR"
   '     .Columns("AUDITOR").TextAlignment = jgexAlignLeft
   '     .Columns("AUDITOR").Width = 2500
        
    '    .Columns("Fec_Auditoria").Caption = "FEC.AUDIT."
    '    .Columns("Fec_Auditoria").TextAlignment = jgexAlignCenter
    '    .Columns("Fec_Auditoria").Width = 1000
                
        .Columns("Observacion").Caption = "OBSERVACIONES"
     
      '  Dim oGroup01 As GridEX20.JSGroup
       ' Dim oGroup02 As GridEX20.JSGroup
        
        'Set oGroup01 = .Groups.Add(.Columns("GRUPO_OT").Index, jgexSortAscending)
        'Set oGroup02 = .Groups.Add(.Columns("GRUPO_MAQUINA").Index, jgexSortAscending)
        
        '.BackColorRowGroup = &H8000000F
        '.DefaultGroupMode = jgexDGMExpanded
        '.ForeColorRowGroup = vbBlue
        
       ' Dim colKILOS As JSColumn
       ' Dim colUNIDADES As JSColumn
        
       ' .GroupFooterStyle = jgexTotalsGroupFooter
        
       ' Set colKILOS = .Columns("Kgs_Rollo")
       ' With colKILOS
       '     .AggregateFunction = jgexSum
       '     .TotalRowPrefix = ""
       ' End With
        'Set colUNIDADES = .Columns("Uni_Rollos")
        'With colUNIDADES
        '    .AggregateFunction = jgexSum
        '    .TotalRowPrefix = ""
        'End With
        
        If .RowCount > 0 Then
            .Row = -1
            .Col = .Columns.Count - 1
        End If
    End With
Exit Sub
Fin:
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
End Sub

Private Sub Form_Load()
    BUSCAR
 '   FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp1, Me.Name)
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    On Error GoTo SALTO_ERROR
    
    Select Case ActionName
        Case "ADICIONAR"
             AddRollo "I", sCod_Almacen, sCod_TipMov, sNum_MovStk
             BUSCAR
             gexMovDet.SetFocus
        Case "MASIVO"
            Frm_DetalleTejeduria.Scod_ordtra = Right(RTrim(Scod_ordtra), 4)
            Frm_DetalleTejeduria.sCod_Almacen = sCod_Almacen
            Frm_DetalleTejeduria.sCod_TipMov = sCod_TipMov
            Frm_DetalleTejeduria.sNum_MovStk = sNum_MovStk
            Frm_DetalleTejeduria.Show 1
            BUSCAR
             
        Case "ELIMINAR"
            Dim sTit As String
            
            sTit = "Sistema De Tejeduria"
            If MsgBox("Desea Eliminar este Rollo?", vbQuestion + vbYesNo, sTit) = vbNo Then Exit Sub
            
            If gexMovDet.RowCount = 0 Then Exit Sub
            
            sFlg_Devolucion_Rollos_Tejeduria = DevuelveCampo("select isnull(Flg_Devolucion_Rollos_Tejeduria,'') from LG_tiposmov where cod_tipmov = '" & sCod_TipMov & "'", cConnect)
            sCOD_TIPMOV_ROLLOS_TEJEDURIA = DevuelveCampo("select isnull(COD_TIPMOV_ROLLOS_TEJEDURIA,'') from LG_tiposmov where cod_tipmov = '" & sCod_TipMov & "'", cConnect)
            
            If sCOD_TIPMOV_ROLLOS_TEJEDURIA = "01" Then
                If sFlg_Devolucion_Rollos_Tejeduria = "S" Then
                    MsgBox "Tipo Movimiento no permite eliminación", vbCritical
                    Exit Sub
                End If
                DelDetMov
            End If
            
            If sCOD_TIPMOV_ROLLOS_TEJEDURIA = "02" Then
                  strSQL = "LG_UP_MAN_TX_MOVISTK_DETALLE_ROLLOS_DESPACHO_PARTIDA 'D', '" & sCod_Almacen & _
                                                                                  "', '" & sNum_MovStk & "','" & LTrim(sNum_Secuencia) & "', '" & _
                                                                                  gexMovDet.Value(gexMovDet.Columns("Prefijo_Maquina").Index) & "', '" & RTrim(gexMovDet.Value(gexMovDet.Columns("Codigo_Rollo").Index)) & "', 0, " & _
                                                                                  0 & ", 'N', '" & vusu & "'"
                  ExecuteSQL cConnect, strSQL
                  BUSCAR
            End If
            
            If sCOD_TIPMOV_ROLLOS_TEJEDURIA = "03" Then
                strSQL = "LG_UP_MAN_TX_MOVISTK_DETALLE_ROLLOS_OTROS_MOVS 'D', '" & sCod_Almacen & _
                                                                         "', '" & sNum_MovStk & "','" & LTrim(sNum_Secuencia) & "', '" & _
                                                                         gexMovDet.Value(gexMovDet.Columns("Prefijo_Maquina").Index) & "', '" & RTrim(gexMovDet.Value(gexMovDet.Columns("Codigo_Rollo").Index)) & "', 0, " & _
                                                                         0 & ", 'N', '" & vusu & "' "
                ExecuteSQL cConnect, strSQL
                BUSCAR
            End If
            gexMovDet.SetFocus
            
        Case "SALIR"
            Unload Me
    End Select
    
    Exit Sub
SALTO_ERROR:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub DelDetMov()
On Error GoTo Fin
Dim sTit As String
    If gexMovDet.RowCount = 0 Then Exit Sub
    sTit = "Eliminar Detalle de Movimiento"
    If MsgBox("Desea Eliminar este Movimento?", vbQuestion + vbYesNo, sTit) = vbNo Then Exit Sub
    
    strSQL = "EXEC lg_UP_MAN_TX_MOVISTK_DETALLE_ROLLOS 'D', '" & _
    gexMovDet.Value(gexMovDet.Columns("Cod_Almacen").Index) & "', '" & _
    gexMovDet.Value(gexMovDet.Columns("Num_MovStk").Index) & "', '" & _
    gexMovDet.Value(gexMovDet.Columns("Num_Secuencia").Index) & "', '" & _
    gexMovDet.Value(gexMovDet.Columns("Prefijo_Maquina").Index) & "', '" & _
    gexMovDet.Value(gexMovDet.Columns("Codigo_Rollo").Index) & "', " & _
    gexMovDet.Value(gexMovDet.Columns("Kgs_Rollo").Index) & ", " & _
    gexMovDet.Value(gexMovDet.Columns("Uni_Rollos").Index) & ", '" & _
    gexMovDet.Value(gexMovDet.Columns("observacion").Index) & "'"
    
    ExecuteSQL cConnect, strSQL
    
    BUSCAR
Exit Sub
Fin:
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
End Sub

Public Sub AddRollo(Accion As String, Cod_Almacen, cod_tipmov As String, Num_MovStk As String)
    On Error GoTo SALTO_ERROR:
    
    Dim rstMov As ADODB.Recordset

    strSQL = "SELECT Des_TipMov, Cod_Calidad, Cod_ClaMov, Cod_TipAnx, Tip_PtMp, Flg_Devolucion_Rollos_Tejeduria,COD_TIPMOV_ROLLOS_TEJEDURIA   " & _
             "FROM LG_TIPOSMOV " & _
             "WHERE Cod_TipMov = '" & cod_tipmov & "'"

    Set rstMov = CargarRecordSetDesconectado(strSQL, cConnect)
    If rstMov.RecordCount > 0 Then rstMov.MoveFirst
    
    If rstMov!COD_TIPMOV_ROLLOS_TEJEDURIA = "01" Then
        If IIf(IsNull(rstMov!Flg_Devolucion_Rollos_Tejeduria), "", rstMov!Flg_Devolucion_Rollos_Tejeduria) = "S" Then
            With frmMovRolloDevolucionDet
                '.oParent = Me
                .sAccion = Accion
                .sCod_Almacen = Cod_Almacen
                .sCod_TipMov = cod_tipmov
                .sDes_TipMov = IIf(IsNull(rstMov!DES_TIPMOV), "", rstMov!DES_TIPMOV)
                .sCod_Calidad = IIf(IsNull(rstMov!Cod_Calidad), "", rstMov!Cod_Calidad)
                .sCod_ClaMov = IIf(IsNull(rstMov!Cod_ClaMov), "", rstMov!Cod_ClaMov)
                .sCod_TipAnx = IIf(IsNull(rstMov!Cod_TipAnx), "", rstMov!Cod_TipAnx)
                .sTip_PtMP = IIf(IsNull(rstMov!Tip_PtMp), "", rstMov!Tip_PtMp)
                .sFlg_Devolucion_Rollos_Tejeduria = IIf(IsNull(rstMov!Flg_Devolucion_Rollos_Tejeduria), "", rstMov!Flg_Devolucion_Rollos_Tejeduria) '
                .sNum_MovStk = Num_MovStk
                .sNum_Secuencia = ""
                .Caption = "Det. Movimiento de Devolucion: " & .sNum_MovStk & " - " & .sDes_TipMov
                rstMov.Close
                .HabilitaCantidades
                .LimpiaForm
                .Show vbModal
            End With
        Else
            With frmMovRolloDet
                .sAccion = Accion
                .sCod_Almacen = Cod_Almacen
                .sCod_TipMov = cod_tipmov
                .sDes_TipMov = IIf(IsNull(rstMov!DES_TIPMOV), "", rstMov!DES_TIPMOV)
                .sCod_Calidad = IIf(IsNull(rstMov!Cod_Calidad), "", rstMov!Cod_Calidad)
                .sCod_ClaMov = IIf(IsNull(rstMov!Cod_ClaMov), "", rstMov!Cod_ClaMov)
                .sCod_TipAnx = IIf(IsNull(rstMov!Cod_TipAnx), "", rstMov!Cod_TipAnx)
                '.sTip_PtMP = IIf(IsNull(rstMov!Tip_PtMp), "", rstMov!Tip_PtMp)
                '.sFlg_Devolucion_Rollos_Tejeduria = IIf(IsNull(rstMov!Flg_Devolucion_Rollos_Tejeduria), "", rstMov!Flg_Devolucion_Rollos_Tejeduria)
                .sNum_MovStk = Num_MovStk
                .sNum_Secuencia = ""
                .Caption = "Detalle de Movimiento : " & .sNum_MovStk & " - " & .sDes_TipMov
                rstMov.Close
                .HabilitaCantidades
                .LimpiaForm
                .Show vbModal
            End With
        End If
    Else
        If rstMov!COD_TIPMOV_ROLLOS_TEJEDURIA = "02" Then
            With frmAddDspDet
                .sAccion = Accion
                .sCod_Almacen = Cod_Almacen
                .sCod_TipMov = cod_tipmov
                '.sDes_TipMov = IIf(IsNull(rstMov!DES_TIPMOV), "", rstMov!DES_TIPMOV)
                '.sCod_Calidad = IIf(IsNull(rstMov!Cod_Calidad), "", rstMov!Cod_Calidad)
                '.sCod_ClaMov = IIf(IsNull(rstMov!Cod_ClaMov), "", rstMov!Cod_ClaMov)
                '.sCod_TipAnx = IIf(IsNull(rstMov!Cod_TipAnx), "", rstMov!Cod_TipAnx)
                '.sTip_PtMP = IIf(IsNull(rstMov!Tip_PtMp), "", rstMov!Tip_PtMp)
                '.sFlg_Devolucion_Rollos_Tejeduria = IIf(IsNull(rstMov!Flg_Devolucion_Rollos_Tejeduria), "", rstMov!Flg_Devolucion_Rollos_Tejeduria)
                .sNum_Secuencia = sNum_Secuencia
                .sNum_MovStk = sNum_MovStk
                .Caption = "Detalle de Movimiento : " & .sNum_MovStk & " - " '& .sDes_TipMov
                .StipoStore = rstMov!COD_TIPMOV_ROLLOS_TEJEDURIA
                 rstMov.Close
                .Show vbModal
           End With
        Else
            If rstMov!COD_TIPMOV_ROLLOS_TEJEDURIA = "03" Then
                With frmAddDspDet
                    .sAccion = Accion
                    .sCod_Almacen = Cod_Almacen
                    .sCod_TipMov = cod_tipmov
                    '.sDes_TipMov = IIf(IsNull(rstMov!DES_TIPMOV), "", rstMov!DES_TIPMOV)
                    '.sCod_Calidad = IIf(IsNull(rstMov!Cod_Calidad), "", rstMov!Cod_Calidad)
                    '.sCod_ClaMov = IIf(IsNull(rstMov!Cod_ClaMov), "", rstMov!Cod_ClaMov)
                    '.sCod_TipAnx = IIf(IsNull(rstMov!Cod_TipAnx), "", rstMov!Cod_TipAnx)
                    '.sTip_PtMP = IIf(IsNull(rstMov!Tip_PtMp), "", rstMov!Tip_PtMp)
                    '.sFlg_Devolucion_Rollos_Tejeduria = IIf(IsNull(rstMov!Flg_Devolucion_Rollos_Tejeduria), "", rstMov!Flg_Devolucion_Rollos_Tejeduria)
                    .sNum_Secuencia = sNum_Secuencia
                    .sNum_MovStk = Num_MovStk
                    .Caption = "Detalle de Movimiento : " & .sNum_MovStk & " - " '& .sDes_TipMov
                    .StipoStore = rstMov!COD_TIPMOV_ROLLOS_TEJEDURIA
                     rstMov.Close
                    .Show vbModal
               End With
            End If
        End If
    End If
    Exit Sub
    
SALTO_ERROR:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub


