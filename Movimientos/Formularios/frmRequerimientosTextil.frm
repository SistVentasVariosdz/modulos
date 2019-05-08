VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmRequerimientosTextil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requerimientos Textiles"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   10125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   510
      Left            =   6975
      TabIndex        =   2
      Top             =   4485
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Cancelar"
      Height          =   510
      Left            =   8550
      TabIndex        =   1
      Top             =   4485
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Requerimientos de Grupo Textíl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   60
      TabIndex        =   0
      Top             =   -15
      Width           =   10020
      Begin SSDataWidgets_B.SSDBGrid DGridLista 
         Height          =   4020
         Index           =   2
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   9795
         _Version        =   196617
         DataMode        =   2
         HeadLines       =   2
         Col.Count       =   15
         BackColorOdd    =   12648447
         RowHeight       =   423
         Columns.Count   =   15
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "COD_TELA"
         Columns(0).Name =   "COD_TELA"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "DES_TELA"
         Columns(1).Name =   "DES_TELA"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3519
         Columns(2).Caption=   "Tela"
         Columns(2).Name =   "Tela"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3200
         Columns(3).Visible=   0   'False
         Columns(3).Caption=   "COD_COMB"
         Columns(3).Name =   "COD_COMB"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Visible=   0   'False
         Columns(4).Caption=   "DES_COMB"
         Columns(4).Name =   "DES_COMB"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3519
         Columns(5).Caption=   "Combinación"
         Columns(5).Name =   "Combinacion"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   1244
         Columns(6).Caption=   "Talla"
         Columns(6).Name =   "COD_TALLA"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   1931
         Columns(7).Caption=   "Reque. Tex"
         Columns(7).Name =   "CAN_CONSTEX"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(7).Mask =   "#######.##"
         Columns(8).Width=   1931
         Columns(8).Caption=   "Reque. Cnf"
         Columns(8).Name =   "CAN_CONSCNF"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(8).Mask =   "#######.##"
         Columns(9).Width=   1958
         Columns(9).Caption=   "Comp. Tex"
         Columns(9).Name =   "CAN_COMPTEX"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         Columns(9).Mask =   "#######.##"
         Columns(10).Width=   1931
         Columns(10).Caption=   "Comp. Cnf"
         Columns(10).Name=   "CAN_COMPCNF"
         Columns(10).DataField=   "Column 10"
         Columns(10).DataType=   8
         Columns(10).FieldLen=   256
         Columns(10).Mask=   "#######.##"
         Columns(11).Width=   1931
         Columns(11).Caption=   "Repo. Tex"
         Columns(11).Name=   "CAN_REPOTEX"
         Columns(11).DataField=   "Column 11"
         Columns(11).DataType=   8
         Columns(11).FieldLen=   256
         Columns(11).Mask=   "#######.##"
         Columns(12).Width=   1931
         Columns(12).Caption=   "Repo. Cnf"
         Columns(12).Name=   "CAN_REPOCNF"
         Columns(12).DataField=   "Column 12"
         Columns(12).DataType=   8
         Columns(12).FieldLen=   256
         Columns(12).Mask=   "#######.##"
         Columns(13).Width=   1958
         Columns(13).Caption=   "Medida"
         Columns(13).Name=   "MEDIDA"
         Columns(13).DataField=   "Column 13"
         Columns(13).DataType=   8
         Columns(13).FieldLen=   256
         Columns(14).Width=   3200
         Columns(14).Visible=   0   'False
         Columns(14).Caption=   "Cod_Medida"
         Columns(14).Name=   "Cod_Medida"
         Columns(14).DataField=   "Column 14"
         Columns(14).DataType=   8
         Columns(14).FieldLen=   256
         _ExtentX        =   17277
         _ExtentY        =   7091
         _StockProps     =   79
         Caption         =   "Requerimientos del Grupo Textíl"
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmRequerimientosTextil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public varCod_OrdPro As String
Public varOpcion As Integer
Dim Rs_Lista As New ADODB.Recordset
Dim Strsql As String
'Estas son variables para la impresion
Public varCod_GrupoTex, varDes_GrupoTex As String
Public varFec_ExpUlt, varAbr_Cliente, varNom_Cliente As String
'Variables para la impresion
Public varCadena_Familias As String
Public varCancelImpresion As Integer
Public varCod_Origen As String

Public oParent As Object

'Sub Reporte()
'On Error GoTo hand
'Dim oo As Object
'Dim Ruta As String
'Dim Usu As String
'Dim Rutalogo As String
'
'    Strsql = "SELECT  Ruta_Logo FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA ='" & vemp & "'"
'    Rutalogo = DevuelveCampo(Strsql, cCONNECT)
'
''    Ruta = vRuta & "\Requerimientos-GrupoTex.XLT"
'    Ruta = App.Path & "\Requerimientos-GrupoTex.XLT"
'    Usu = "Usuario : " & vusu
'    Set oo = CreateObject("excel.application")
'    oo.Workbooks.Open Ruta
'    oo.Visible = True
'    oo.DisplayAlerts = False
'
'    oo.Run "Reporte", varOpcion, varCadena_Familias, varCod_GrupoTex & " " & varDes_GrupoTex, CStr(varFec_ExpUlt), DEVUELVE_OPS, varAbr_Cliente & " " & varNom_Cliente, Rutalogo, cCONNECT
'    Set oo = Nothing
'Exit Sub
'hand:
'    ErrorHandler Err, "GeneraReportes"
'    Set oo = Nothing
'
'End Sub
'Private Function DEVUELVE_OPS() As String
'    Dim Rs_Lista1 As New ADODB.Recordset
'    Dim Result As String
'    On Error GoTo Cargar_DatosErr
'
'    Strsql = "SELECT cod_ordpro FROM ES_ORDPRO WHERE Cod_GrupoTex ='" & varCod_OrdPro & "'"
'
'    'Strsql = "EXEC UP_SEL_TOTALORDPROREQ_TEXTIL '" & varCod_GrupoTex & "'," & CStr(opcion)
'
'    Set Rs_Lista1 = Nothing
'    Rs_Lista1.ActiveConnection = cCONNECT
'    Rs_Lista1.CursorType = adOpenStatic
'    Rs_Lista1.CursorLocation = adUseClient
'    Rs_Lista1.LockType = adLockReadOnly
'    Rs_Lista1.Open Strsql
'
'    Result = ""
'
'    If Rs_Lista1.RecordCount > 0 Then
'        Rs_Lista1.MoveFirst
'        While Not Rs_Lista1.EOF And Not Rs_Lista1.BOF
'            Result = Result & "  " & Trim(Rs_Lista1(0).Value)
'            Rs_Lista1.MoveNext
'        Wend
'        Rs_Lista1.Close
'        Set Rs_Lista1 = Nothing
'    End If
'
'    DEVUELVE_OPS = Result
'    Exit Function
'Cargar_DatosErr:
'    Set Rs_Lista = Nothing
'    ErrorHandler Err, "Cargar_Datos"
'
'End Function

Public Sub CARGA_LISTA(opcion As Integer)
    On Error GoTo Cargar_DatosErr
    Dim Rs_Prov As New ADODB.Recordset
    
    Strsql = "EXEC UP_SEL_TOTALORDPROREQ_TEXTIL '" & varCod_GrupoTex & "'," & CStr(opcion)
    
    Set Rs_Lista = Nothing
    Rs_Lista.ActiveConnection = cConnect
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    Rs_Lista.Open Strsql

    Set Rs_Prov = Rs_Lista.Clone

   'Esto es para asignar la data al grid
        
        Select Case opcion
            Case 1
                    Me.DGridLista(0).Redraw = False
                    SSDBGridSetGrid Me.DGridLista(0)
                    ADODBToSSDBGrid Rs_Prov, DGridLista(0)
                    DGridLista(0).ActiveRowStyleSet = "RowActive"
                    DGridLista(0).SelectTypeRow = ssSelectionTypeMultiSelectRange
                    DGridLista(0).Visible = True
            Case 2
                    Me.DGridLista(1).Redraw = False
                    SSDBGridSetGrid Me.DGridLista(1)
                    ADODBToSSDBGrid Rs_Prov, DGridLista(1)
                    DGridLista(1).ActiveRowStyleSet = "RowActive"
                    DGridLista(1).SelectTypeRow = ssSelectionTypeMultiSelectRange
                    DGridLista(1).Visible = True
            Case 3
                    Me.DGridLista(2).Redraw = False
                    SSDBGridSetGrid Me.DGridLista(2)
                    ADODBToSSDBGrid Rs_Prov, DGridLista(2)
                    DGridLista(2).ActiveRowStyleSet = "RowActive"
                    DGridLista(2).SelectTypeRow = ssSelectionTypeMultiSelectRange
                    DGridLista(2).Visible = True
            Case 4
                    Me.DGridLista(3).Redraw = False
                    SSDBGridSetGrid Me.DGridLista(3)
                    ADODBToSSDBGrid Rs_Prov, DGridLista(3)
                    DGridLista(3).ActiveRowStyleSet = "RowActive"
                    DGridLista(3).SelectTypeRow = ssSelectionTypeMultiSelectRange
                    DGridLista(3).Visible = True
        End Select
       
   'Aqui termina la asignación de la data



    'Set DGridLista.DataSource = Rs_Lista
    'DGridLista_RowColChange 0, 0
    'If Rs_Lista.RecordCount > 0 Then
    '    DGridLista.Enabled = True
    '    HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    'Else
    '    HabilitaMant Me.MantFunc1, "ADICIONAR"
    'End If
    
    'Call FORMATEA_GRID(opcion)
    
    Rs_Prov.Close
    Set Rs_Prov = Nothing
    
    Exit Sub
Cargar_DatosErr:
    Set Rs_Lista = Nothing
    ErrorHandler err, "Cargar_Datos"
End Sub

Private Sub cmdAceptar_Click()
    Select Case varOpcion
        Case 1
                oParent.TxtTela = Me.DGridLista(0).Columns("Hilo").Value
        Case 2
                oParent.TxtTela = Me.DGridLista(1).Columns("Hilo").Value
        Case 3
                oParent.TxtTela = Me.DGridLista(2).Columns("Tela").Value
        Case 4
                oParent.TxtTela = Me.DGridLista(3).Columns("Tela").Value
    End Select
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub DGridLista_DblClick(Index As Integer)
    Select Case Index
        Case 2
                cmdAceptar_Click
        Case 3
                cmdAceptar_Click
    End Select
End Sub


Private Sub Form_Load()

    'Me.FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)

    SSDBGridSetGrid0 Me.DGridLista(0)
    SSDBGridSetGrid0 Me.DGridLista(1)
    SSDBGridSetGrid0 Me.DGridLista(2)
    SSDBGridSetGrid0 Me.DGridLista(3)
End Sub

'Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
'    If Rs_Lista.RecordCount = 0 Then
'        MsgBox "Debe seleccionar un registro para acceder a esta opción. Sirvase Verificar", vbInformation, "Requerimientos"
'        Exit Sub
'    End If
'    Select Case ActionName
'        Case "IMPRIMIR"
'
'            If varOpcion = 3 Or varOpcion = 4 Then
'
'                Load frmSelecFamilias
'                If varOpcion = 3 Then
'                    frmSelecFamilias.varTipoReq = "TC"   'Esto implica q es Tela Cruda
'                Else
'                    frmSelecFamilias.varTipoReq = "TT"   'Esto implica q es Tela Teñida
'                End If
'                frmSelecFamilias.varCod_Grupo = varCod_GrupoTex
'                frmSelecFamilias.CARGA_FAMILIAS
'                frmSelecFamilias.Frame1.Visible = False
'                Set frmSelecFamilias.oParent = Me
'                frmSelecFamilias.Show 1
'
'                If Me.varCancelImpresion = 1 Then
'                     Exit Sub
'                End If
'
'            End If
'
'            'Estas son las varables de la cabecera
'            Reporte
'        Case "DETALLE"
'
'            Load frmRequerimientosTextilDet
'
'            'Call frmRequerimientosTextilDet.FORMATEA_GRID(varOpcion)
'
'            Select Case varOpcion
'                Case 1
'                        frmRequerimientosTextilDet.varCod_HilTel = DGridLista(0).Columns("Cod_HilTel").Value
'                        frmRequerimientosTextilDet.varCod_GrupoTex = varCod_GrupoTex
'                Case 2
'                        frmRequerimientosTextilDet.varCod_HilTel = DGridLista(1).Columns("Cod_HilTel").Value
'                        frmRequerimientosTextilDet.varCod_GrupoTex = varCod_GrupoTex
'                        frmRequerimientosTextilDet.varCod_Color = DGridLista(1).Columns("Cod_Color").Value
'                Case 3
'                        frmRequerimientosTextilDet.varCod_GrupoTex = varCod_GrupoTex
'                        frmRequerimientosTextilDet.varCod_Tela = DGridLista(2).Columns("Cod_Tela").Value
'                        frmRequerimientosTextilDet.varCod_Comb = DGridLista(2).Columns("Cod_Comb").Value
'                        frmRequerimientosTextilDet.varCod_Talla = DGridLista(2).Columns("Cod_Talla").Value
'                        frmRequerimientosTextilDet.varCod_Medida = DGridLista(2).Columns("Cod_Medida").Value
'                Case 4
'                        frmRequerimientosTextilDet.varCod_GrupoTex = varCod_GrupoTex
'                        frmRequerimientosTextilDet.varCod_Color = DGridLista(3).Columns("Cod_Color").Value
'                        frmRequerimientosTextilDet.varCod_Tela = DGridLista(3).Columns("Cod_Tela").Value
'                        frmRequerimientosTextilDet.varCod_Comb = DGridLista(3).Columns("Cod_Comb").Value
'                        frmRequerimientosTextilDet.varCod_Talla = DGridLista(3).Columns("Cod_Talla").Value
'                        frmRequerimientosTextilDet.varCod_Medida = DGridLista(3).Columns("Cod_Medida").Value
'            End Select
'
'            Call frmRequerimientosTextilDet.CARGA_LISTA(varOpcion)
'            frmRequerimientosTextilDet.Show 1
'        Case "PROGRAMACION"
'            Load frmProgramacion
'            frmProgramacion.varOpcion = Me.varOpcion
'            Select Case varOpcion
'                Case 1
'                        frmProgramacion.varCod_Tela = DGridLista(0).Columns("Cod_HilTel").Value
'                        frmProgramacion.varCod_GrupoTex = varCod_GrupoTex
'
'                        frmProgramacion.Frame1.Visible = True
'                        frmProgramacion.txtGrupo1.Text = varCod_GrupoTex
'                        frmProgramacion.txtHilo1.Text = DGridLista(0).Columns("Hilo").Value
'                Case 2
'                        frmProgramacion.varCod_Tela = DGridLista(1).Columns("Cod_HilTel").Value
'                        frmProgramacion.varCod_GrupoTex = varCod_GrupoTex
'                        frmProgramacion.varCod_Color = DGridLista(1).Columns("Cod_Color").Value
'
'                        frmProgramacion.Frame2.Visible = True
'                        frmProgramacion.txtGrupo2.Text = varCod_GrupoTex
'                        frmProgramacion.txtHilo2.Text = DGridLista(1).Columns("Hilo").Value
'                        frmProgramacion.txtColor2.Text = DGridLista(1).Columns("Color").Value
'
'                Case 3
'                        frmProgramacion.varCod_GrupoTex = varCod_GrupoTex
'                        frmProgramacion.varCod_Tela = DGridLista(2).Columns("Cod_Tela").Value
'                        frmProgramacion.varCod_Comb = DGridLista(2).Columns("Cod_Comb").Value
'                        frmProgramacion.varCod_Talla = DGridLista(2).Columns("Cod_Talla").Value
'
'                        frmProgramacion.Frame3.Visible = True
'                        frmProgramacion.txtGrupo3.Text = varCod_GrupoTex
'                        frmProgramacion.txtTela3.Text = DGridLista(2).Columns("Tela").Value
'                        frmProgramacion.txtComb3.Text = DGridLista(2).Columns("Combinacion").Value
'                        frmProgramacion.txtTalla3.Text = DGridLista(2).Columns("Cod_Talla").Value
'                Case 4
'                        frmProgramacion.varCod_GrupoTex = varCod_GrupoTex
'                        frmProgramacion.varCod_Color = DGridLista(3).Columns("Cod_Color").Value
'                        frmProgramacion.varCod_Tela = DGridLista(3).Columns("Cod_Tela").Value
'                        frmProgramacion.varCod_Comb = DGridLista(3).Columns("Cod_Comb").Value
'                        frmProgramacion.varCod_Talla = DGridLista(3).Columns("Cod_Talla").Value
'
'                        frmProgramacion.Frame4.Visible = True
'                        frmProgramacion.txtGrupo4.Text = varCod_GrupoTex
'                        frmProgramacion.txtTela4.Text = DGridLista(3).Columns("Tela").Value
'                        frmProgramacion.txtComb4.Text = DGridLista(3).Columns("Combinacion").Value
'                        frmProgramacion.txtColor4.Text = DGridLista(3).Columns("Color").Value
'                        frmProgramacion.txtTalla4.Text = DGridLista(3).Columns("Cod_Talla").Value
'            End Select
'            frmProgramacion.CARGA_GRID
'            frmProgramacion.Show 1
'        Case "DETALLEOC"
'            Load frmRequerimientosTextilDetOC
'
'            With frmRequerimientosTextilDetOC
'                Select Case varOpcion
'                    Case 1
'                            .varCod_HilTel = DGridLista(0).Columns("Cod_HilTel").Value
'                            .varCod_GrupoTex = varCod_GrupoTex
'                    Case 2
'                            .varCod_HilTel = DGridLista(1).Columns("Cod_HilTel").Value
'                            .varCod_GrupoTex = varCod_GrupoTex
'                            .varCod_Color = DGridLista(1).Columns("Cod_Color").Value
'                    Case 3
'                            .varCod_GrupoTex = varCod_GrupoTex
'                            .varCod_Tela = DGridLista(2).Columns("Cod_Tela").Value
'                            .varCod_Comb = DGridLista(2).Columns("Cod_Comb").Value
'                            .varCod_Talla = DGridLista(2).Columns("Cod_Talla").Value
'                            .varCod_Medida = DGridLista(2).Columns("Cod_Medida").Value
'                    Case 4
'                            .varCod_GrupoTex = varCod_GrupoTex
'                            .varCod_Color = DGridLista(3).Columns("Cod_Color").Value
'                            .varCod_Tela = DGridLista(3).Columns("Cod_Tela").Value
'                            .varCod_Comb = DGridLista(3).Columns("Cod_Comb").Value
'                            .varCod_Talla = DGridLista(3).Columns("Cod_Talla").Value
'                            .varCod_Medida = DGridLista(3).Columns("Cod_Medida").Value
'                End Select
'            End With
'            Call frmRequerimientosTextilDetOC.CARGA_LISTA(varOpcion)
'            frmRequerimientosTextilDetOC.Show 1
'
'
'    End Select
'End Sub

