VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmMotivo_Notas 
   Caption         =   "Motivos Notas"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraBuscar 
      Height          =   915
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10095
      Begin VB.TextBox txtDes_TipDoc 
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   1905
      End
      Begin VB.TextBox txtCod_TipDoc 
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   0
         Top             =   240
         Width           =   480
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   435
         Left            =   8760
         TabIndex        =   3
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Doc :"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   255
         Width           =   855
      End
   End
   Begin GridEX20.GridEX grxData 
      Height          =   4365
      Left            =   30
      TabIndex        =   1
      Top             =   960
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   7699
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "FrmMotivo_Notas.frx":0000
      Column(2)       =   "FrmMotivo_Notas.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "FrmMotivo_Notas.frx":016C
      FormatStyle(2)  =   "FrmMotivo_Notas.frx":02A4
      FormatStyle(3)  =   "FrmMotivo_Notas.frx":0354
      FormatStyle(4)  =   "FrmMotivo_Notas.frx":0408
      FormatStyle(5)  =   "FrmMotivo_Notas.frx":04E0
      FormatStyle(6)  =   "FrmMotivo_Notas.frx":0598
      ImageCount      =   0
      PrinterProperties=   "FrmMotivo_Notas.frx":0678
   End
   Begin FunctionsButtons.FunctButt fnbOperacion 
      Height          =   510
      Left            =   2595
      TabIndex        =   4
      Top             =   5370
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   900
      Custom          =   $"FrmMotivo_Notas.frx":0850
      Orientacion     =   0
      Style           =   0
      Language        =   1
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   240
      Top             =   5400
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmMotivo_Notas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public codigo As String, Descripcion As String
Dim strSQL As String


Private Sub cmdBuscar_Click()
    Dim adoRs As ADODB.Recordset

    Dim strOpcion As String
    strSQL = "EXEC VENTAS_MUESTRA_MOTIVOS_NOTAS '" & txtCod_TipDoc.Text & "'"
    Set adoRs = CargarRecordSetDesconectado(strSQL, cCONNECT)
    Set grxData.ADORecordset = adoRs
    Call CONFIGURAR_GRID
End Sub

Public Sub CONFIGURAR_GRID()
    grxData.Columns("Cod_TipDoc").Width = "0"
    grxData.Columns("Cod_Mot_Nota").Width = "600"
    grxData.Columns("Descripcion").Width = "2200"
    grxData.Columns("Cuenta").Width = "0"
    grxData.Columns("Des_CtaCont").Width = "2500"
    
    grxData.Columns("Flg_Mostrar_Grupo").Width = "1000"
    grxData.Columns("Flg_No_Cantidad_Gupo").Width = "1000"
    grxData.Columns("Flg_Gasto_Financiero").Width = "1000"
    grxData.Columns("Flg_Condonacion_Deuda").Width = "1000"
End Sub

Private Sub fnbOperacion_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
  Dim ELIMINAR As Integer
Select Case ActionName
Case "ADICIONAR"
    FrmDetalleMotivoNotas.sOpcion = "I"
    FrmDetalleMotivoNotas.Sid_proyeccion = 0
    FrmDetalleMotivoNotas.Show 1
    cmdBuscar_Click
Case "MODIFICAR"
    FrmDetalleMotivoNotas.sOpcion = "U"
    FrmDetalleMotivoNotas.Sid_proyeccion = grxData.Value(grxData.Columns("Cod_Mot_Nota").Index)
    FrmDetalleMotivoNotas.txtCod_TipDoc = UCase(grxData.Value(grxData.Columns("Cod_TipDoc").Index))
    FrmDetalleMotivoNotas.txtDes_TipDoc = DevuelveCampo("SELECT Des_TipDoc FROM CN_TiposDocum WHERE Cod_TipDoc = '" & grxData.Value(grxData.Columns("Cod_TipDoc").Index) & "'", cCONNECT)
    FrmDetalleMotivoNotas.txtdescripcion1 = grxData.Value(grxData.Columns("Descripcion").Index)
    FrmDetalleMotivoNotas.txtCuenta = grxData.Value(grxData.Columns("Cuenta").Index)
    FrmDetalleMotivoNotas.txtDescripcion = grxData.Value(grxData.Columns("Des_CtaCont").Index)
    FrmDetalleMotivoNotas.txtCuenta2010 = grxData.Value(grxData.Columns("cod_ctacont_hasta_2010").Index)
    FrmDetalleMotivoNotas.txtDescripcion2010 = grxData.Value(grxData.Columns("des_ctacont_hasta_2010").Index)
    
    If Trim(grxData.Value(grxData.Columns("Flg_Gasto_Financiero").Index)) = "S" Then
        FrmDetalleMotivoNotas.chkGasto_Financiero.Value = 1
    Else
        FrmDetalleMotivoNotas.chkGasto_Financiero.Value = 0
    End If
    
    FrmDetalleMotivoNotas.txtCod_TipDoc.Enabled = False
    FrmDetalleMotivoNotas.txtDes_TipDoc.Enabled = False
    
    FrmDetalleMotivoNotas.Show 1
    cmdBuscar_Click
Case "ELIMINAR"
            ELIMINAR = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Motivo Notas")
            If ELIMINAR = vbYes Then
                Call Eliminar_Datos
                cmdBuscar_Click
            End If
Case "IMPRIMIR"
'    Call Reporte
Case "SALIR"
    Unload Me
End Select
End Sub


Sub Eliminar_Datos()
    Dim strSQL As String
    On Error GoTo Salvar_DatosErr

 
    strSQL = "EXEC VENTAS_MAN_MOTIVOS_NOTAS 'D','" & UCase(grxData.Value(grxData.Columns("Cod_TipDoc").Index)) & "','" & Trim(grxData.Value(grxData.Columns("Cod_Mot_Nota").Index)) & "','','','',''"
      
    ExecuteCommandSQL cCONNECT, strSQL

    MsgBox "Registro eliminado satisfactoriamente......", vbInformation, Me.Caption
    
    Exit Sub
Salvar_DatosErr:
    ErrorHandler err, "Salvar_Datos"
End Sub

Private Sub txtCod_TipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_TipDoc.Text) = "" Then
            Call BUSCA_tipo(3)
        Else
            Call BUSCA_tipo(1)
        End If
    End If
End Sub

Public Sub BUSCA_tipo(Tipo As Integer)
    Select Case Tipo
        Case 1:
                    strSQL = "select Des_TipDoc as Descripcion from CN_TiposDocum  where  Flg_Doc_Ventas = '*' and Cod_TipDoc='" & txtCod_TipDoc.Text & "'"
                    txtDes_TipDoc.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                    cmdBuscar.SetFocus
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim RS As Object
                    Set RS = CreateObject("ADODB.Recordset")
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "select Cod_TipDoc as codigo, Des_TipDoc as Descripcion  from CN_TiposDocum  where  Flg_Doc_Ventas = '*' and Des_TipDoc like '%" & txtDes_TipDoc.Text & "%'"
                    Else
                        oTipo.SQuery = "select Cod_TipDoc as Codigo , Des_TipDoc as Descripcion from CN_TiposDocum  where  Flg_Doc_Ventas = '*'"
                    End If
                    
                    oTipo.CARGAR_DATOS
                   ' oTipo.DGridLista.Columns(2).Width = 2500
                    oTipo.Show 1
                    If codigo <> "" Then
                         txtCod_TipDoc.Text = Trim(codigo)
                         txtDes_TipDoc.Text = Trim(Descripcion)

                         codigo = "": Descripcion = ""
                        cmdBuscar.SetFocus
                    End If
                    Set oTipo = Nothing
                    Set RS = Nothing
    End Select
    
End Sub


Private Sub txtDes_TipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDes_TipDoc.Text) = "" Then
            Call BUSCA_tipo(3)
        Else
            Call BUSCA_tipo(2)
        End If
    End If

End Sub
