VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSeleccionCotizacion 
   Caption         =   "Seleccione Cotizacion:"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   3540
      TabIndex        =   7
      Top             =   4260
      Width           =   1665
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Siguiente >"
      Height          =   555
      Left            =   1230
      TabIndex        =   6
      Top             =   4230
      Width           =   1635
   End
   Begin VB.CommandButton CmdDelAll 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3420
      TabIndex        =   4
      Top             =   2880
      Width           =   525
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3420
      TabIndex        =   3
      Top             =   2310
      Width           =   525
   End
   Begin VB.CommandButton CmdAddAll 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3420
      TabIndex        =   2
      Top             =   1470
      Width           =   525
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3420
      TabIndex        =   1
      Top             =   900
      Width           =   525
   End
   Begin MSDataGridLib.DataGrid DG_Origen 
      Height          =   3225
      Left            =   240
      TabIndex        =   0
      Top             =   540
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5689
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Origen"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "COD_ESTCLI"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "DES_ESTCLI"
         Caption         =   "Descripción"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1500.095
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DG_Destino 
      Height          =   3255
      Left            =   4200
      TabIndex        =   5
      Top             =   570
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   5741
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Destino"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "COD_ESTCLI"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "DES_ESTCLI"
         Caption         =   "Descripción"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2009.764
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSeleccionCotizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOrigen As New ADODB.Recordset
Dim rsDestino As New ADODB.Recordset
Dim varNumCot As Integer
Public varCod_Cliente, varCod_TemCli As String
Public varEst_Avance As Boolean
Public varFSolicitud As Date
Public varFEntrega As Date
Public varFEntProto As Date
Public varObs As String
Public varCod_EstPro As String

Public varAbr_Cliente As String
Public varDes_Cliente As String
Public varNom_TemCli As String
'Public varNumCot As Integer
Public varCod_EstCli As String


Public Sub CARGA_DATA()
Dim Rs As New ADODB.Recordset
Dim srtSql As String
    Rs.ActiveConnection = cCONNECT
    Rs.CursorType = adOpenStatic
    Rs.CursorLocation = adUseClient
    Rs.LockType = adLockOptimistic

    'Extrae los estilos que no tienen cotizaciones
    srtSql = "EXEC SM_EXTRAE_ESTCLI_SINCOT '" & varCod_Cliente & "','" & varCod_TemCli & "'"
    Rs.Open srtSql
    Set rsOrigen = grsCopy(Rs)
    Set DG_Origen.DataSource = rsOrigen
    EstadoBotones 0, False
    
    If rsOrigen.RecordCount > 0 Then
        CmdAdd.Enabled = True
        CmdAddAll.Enabled = True
    End If
    
   'Creamos los campos la grilla destino
    rsDestino.Fields.Append rsOrigen.Fields(0).Name, rsOrigen.Fields(0).Type, rsOrigen.Fields(0).DefinedSize
    rsDestino.Fields.Append rsOrigen.Fields(1).Name, rsOrigen.Fields(1).Type, rsOrigen.Fields(1).DefinedSize
    rsDestino.Open
    
     Set DG_Destino.DataSource = rsDestino

End Sub

Private Sub EstadoBotones(Tipo As Integer, Valor As Boolean)
Select Case Tipo
    Case 0
        CmdAdd.Enabled = Valor
        CmdAddAll.Enabled = Valor
        CmdDel.Enabled = Valor
        CmdDelAll.Enabled = Valor
    Case 1
        CmdAdd.Enabled = Valor
        CmdAddAll.Enabled = Valor
        CmdDel.Enabled = Not Valor
        CmdDelAll.Enabled = Not Valor
End Select
End Sub

Private Sub cmdAceptar_Click()
Dim cn As New ADODB.Connection
Dim CMD As New ADODB.Command
Dim CMD2 As New ADODB.Command
Dim PM As ADODB.Parameter
Dim sMessage As Integer
'Dim Rs As New ADODB.Recordset

cn.Open cCONNECT
If Not (rsDestino.EOF And rsDestino.BOF) Then
    varEst_Avance = True
    frmSeleccionCotizacionSgt.Show 1
    
    If varEst_Avance Then
        'Obtiene el valor maximo +1 de la tabla Tg_Control
        With CMD2
            Set .ActiveConnection = cn
            .CommandType = adCmdStoredProc
            .CommandText = "UP_CONTROL_SOLICITUD_CONS"
        End With
        Set PM = CMD2.CreateParameter("@num_solicitud_cons", adInteger, adParamOutput)
        CMD2.Parameters.Append PM
        PM.Value = 0
        CMD2.Execute
        varNumCot = PM.Value
                
           With CMD
                Set .ActiveConnection = cn
                .CommandType = adCmdStoredProc
                .CommandText = "UP_GENERA_COTIZACION"
                
                'Pasa los parametros para el procedimiento UP_GENERA_COTIZACION
                .Parameters.Append .CreateParameter("@NUM_COTIZACION", adInteger, adParamInput, , varNumCot)
                .Parameters.Append .CreateParameter("@COD_CLIENTE", adVarChar, adParamInput, 5, varCod_Cliente)
                .Parameters.Append .CreateParameter("@COD_TEMCLI", adVarChar, adParamInput, 3, varCod_TemCli)
                .Parameters.Append .CreateParameter("@COD_ESTCLI", adVarChar, adParamInput, 20, "")
                .Parameters.Append .CreateParameter("@F_SOLICITUD", adDate, adParamInput, , varFSolicitud)
                .Parameters.Append .CreateParameter("@F_ENTREGA", adDate, adParamInput, 3, varFEntrega)
                .Parameters.Append .CreateParameter("@F_ENTREGA_PROTO", adDate, adParamInput, 3, varFEntProto)
                .Parameters.Append .CreateParameter("@OBS", adVarChar, adParamInput, 250, varObs)
                .Parameters.Append .CreateParameter("@COD_USUARIO", adVarChar, adParamInput, 15, vusu)
            End With
            
            rsDestino.MoveFirst
            With rsDestino
                Do Until .EOF
                    CMD.Parameters("@COD_ESTCLI").Value = .Fields("COD_ESTCLI").Value
                    CMD.Execute , , adExecuteNoRecords
                    .MoveNext
                Loop
            End With
            
'            If varEst_Cot Then
'                CARGA_ESTCLI
                sMessage = MsgBox("Se generó satisfactoriamente la Cotizacion Nº " & varNumCot & ".  ¿Desea imprimirla ahora?", vbYesNo, "Generación de Cotizacion")
                Select Case sMessage
                    Case vbYes
                        Load frmRepEstCliTem
                        frmRepEstCliTem.varCod_EstPro = Me.varCod_EstPro
                        'frmRepEstCliTem.frPrincipal.Visible = False
                        'frmRepEstCliTem.frFecha.Visible = True
                        'frmRepEstCliTem.frFecha.Top = 0
                        'frmRepEstCliTem.OptEstiloSeleccionado.Value = True
                        frmRepEstCliTem.varNumCot = varNumCot
                        
                        frmRepEstCliTem.varAbr_Cliente = Me.varAbr_Cliente
                        frmRepEstCliTem.varCod_TemCli = Me.varCod_TemCli
                        frmRepEstCliTem.varDes_Cliente = Me.varDes_Cliente
                        frmRepEstCliTem.varNom_TemCli = Me.varNom_TemCli
                        frmRepEstCliTem.varCod_EstCli = Me.varCod_EstCli
                        frmRepEstCliTem.varObs = Me.varObs
                        frmRepEstCliTem.Show vbModal
                        Set frmRepEstCliTem = Nothing
                    Case vbNo
'                        Exit Sub
                End Select
'            End If

'            frmEstCliTem.varEst_Cot = True
'            frmEstCliTem.varNumCot = varNumCot
'            frmEstCliTem.varObs = Me.varObs
    End If
Else
        MsgBox "No ha seleccionado ningun registro", vbInformation, Me.Caption
        frmEstCliTem.varEst_Cot = False
        Exit Sub
End If
cn.Close
Set cn = Nothing
cmdCancelar_Click
End Sub

Sub GeneraReportes()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String
Dim StrSQL As String
Dim strSQL2 As String

    StrSQL = "select tip_fabrica from tg_control"
    If DevuelveCampo(StrSQL, cCONNECT) = 1 Then
    '    Ruta = "C:\Archivos de programa\Gestion de pedidos\\prototipo.xlt"
        Ruta = vRuta & "\prototipo.xlt"
    Else
    '    Ruta = "C:\Archivos de programa\Gestion de pedidos\prototipoD.xlt"
        Ruta = vRuta & "\prototipoD.xlt"
    End If

    StrSQL = "SELECT NOM_Cliente FROM TG_CLIENTE WHERE cod_Cliente='" & Trim(varCod_Cliente) & "'"
    strSQL2 = "SELECT NOM_TEMCLI FROM TG_TEMCLI WHERE COD_CLIENTE='" & Trim(varCod_Cliente) & "' AND COD_TEMCLI='" & Trim(varCod_TemCli) & "'"
    Set oo = CreateObject("excel.application")
    oo.workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    'oo.run "Reporte", CStr(DevuelveCampo(strSQL, cCONNECT)), frmEstCliTem.txtCod_TemCli, varNumCot, cCONNECT, vemp, frmEstCliTem.txtDes_Cliente, frmEstCliTem.txtNom_TemCli, varObs, vusu
    oo.run "Reporte", CStr(varCod_Cliente), CStr(varCod_TemCli), varNumCot, cCONNECT, vemp, CStr(DevuelveCampo(StrSQL, cCONNECT)), CStr(DevuelveCampo(strSQL2, cCONNECT)), varObs, vusu
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler Err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub CmdAdd_Click()
    
    rsDestino.AddNew
    rsDestino.Fields("COD_ESTCLI").Value = rsOrigen.Fields("COD_ESTCLI").Value
    rsDestino.Fields("DES_ESTCLI").Value = rsOrigen.Fields("DES_ESTCLI").Value

    rsDestino.Update
        
    rsOrigen.Delete
    
    If (rsOrigen.RecordCount > 0 And rsDestino.RecordCount = 0) Then EstadoBotones 1, True
    If (rsDestino.RecordCount > 0 And rsOrigen.RecordCount > 0) Then EstadoBotones 0, True
    
    If rsOrigen.RecordCount = 0 Then EstadoBotones 1, False
End Sub

Private Sub CmdAddAll_Click()
  
 rsOrigen.MoveFirst
 If rsOrigen.RecordCount > 0 Then
    While Not rsOrigen.EOF
        rsDestino.AddNew
        rsDestino.Fields("COD_ESTCLI").Value = rsOrigen.Fields("COD_ESTCLI").Value
        rsDestino.Fields("DES_ESTCLI").Value = rsOrigen.Fields("DES_ESTCLI").Value
        rsDestino.Update
        rsOrigen.Delete
        rsOrigen.MoveNext
    Wend
 End If

    EstadoBotones 1, False
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdDel_Click()
    rsOrigen.AddNew
    rsOrigen.Fields("COD_ESTCLI").Value = rsDestino.Fields("COD_ESTCLI").Value
    rsOrigen.Fields("DES_ESTCLI").Value = rsDestino.Fields("DES_ESTCLI").Value

    rsOrigen.Update
        
    rsDestino.Delete
    
    If (rsDestino.RecordCount > 0 And rsOrigen.RecordCount > 0) Then EstadoBotones 0, True
    If (rsOrigen.RecordCount > 0 And rsDestino.RecordCount = 0) Then EstadoBotones 1, False
    If rsDestino.RecordCount = 0 Then EstadoBotones 1, True

End Sub

Private Sub CmdDelAll_Click()

 rsDestino.MoveFirst
 If rsDestino.RecordCount > 0 Then
    While Not rsDestino.EOF
        rsOrigen.AddNew
        rsOrigen.Fields("COD_ESTCLI").Value = rsDestino.Fields("COD_ESTCLI").Value
        rsOrigen.Fields("DES_ESTCLI").Value = rsDestino.Fields("DES_ESTCLI").Value
        rsOrigen.Update
        rsDestino.Delete
        rsDestino.MoveNext
    Wend
 End If

    EstadoBotones 1, True

End Sub

Private Function grsCopy(ByRef SourceRecordset As ADODB.Recordset, Optional ByVal Records As Long = -1) As ADODB.Recordset

  Dim Rs As ADODB.Recordset
  Dim fld As ADODB.FIELD

  Set Rs = New Recordset
  With Rs
'+++ copio los campos
    For Each fld In SourceRecordset.Fields
      .Fields.Append fld.Name, fld.Type, fld.DefinedSize, fld.Attributes
      .Fields(fld.Name).NumericScale = fld.NumericScale
      .Fields(fld.Name).Precision = fld.Precision
    Next
'--- copio los campos

'+++ copio los valores
    .Open
    While (Not SourceRecordset.EOF) And (Records <> 0)
      .AddNew
      For Each fld In SourceRecordset.Fields
        .Fields(fld.Name).Value = fld.Value
      Next
      .Update
      SourceRecordset.MoveNext
      Records = Records - 1
    Wend
'--- copio los valores
    
'+++ me muevo al inicio del recordset
    If Not (Rs.EOF And Rs.BOF) Then Rs.MoveFirst
'--- me muevo al inicio del recordset

  End With
  Set grsCopy = Rs

End Function

Private Sub Form_Unload(Cancel As Integer)
    Set rsOrigen = Nothing
    Set rsDestino = Nothing
End Sub
