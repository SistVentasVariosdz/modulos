VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAuditoriaCalidadRecepcionAvios 
   Caption         =   "Auditoria Calidad Recepcion Avios"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6375
      Begin VB.TextBox TxtObservaciones 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   2520
         Width           =   3600
      End
      Begin VB.TextBox TxtPorcDesaprobacion 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   7
         Top             =   1800
         Width           =   840
      End
      Begin VB.TextBox TxtAql 
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   2160
         Width           =   840
      End
      Begin VB.TextBox TxtCantidadDesaprobada 
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   1800
         Width           =   840
      End
      Begin VB.TextBox TxtDes_Motivo 
         Height          =   300
         Left            =   3000
         TabIndex        =   4
         Top             =   1440
         Width           =   2685
      End
      Begin VB.TextBox TxtCod_Motivo 
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   1440
         Width           =   675
      End
      Begin VB.OptionButton optDesaprobado 
         Caption         =   "Desaprobado"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton optAprobado 
         Caption         =   "Aprobado"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   2280
         TabIndex        =   0
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Format          =   64094209
         CurrentDate     =   39196
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   840
         TabIndex        =   16
         Top             =   2595
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "% Desaprobacion"
         Height          =   195
         Left            =   3240
         TabIndex        =   15
         Top             =   1875
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "AQL"
         Height          =   195
         Left            =   840
         TabIndex        =   14
         Top             =   2235
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cant. Desaprobada"
         Height          =   195
         Left            =   840
         TabIndex        =   13
         Top             =   1875
         Width           =   1380
      End
      Begin VB.Label Label7 
         Caption         =   "Motivo"
         Height          =   225
         Left            =   840
         TabIndex        =   12
         Top             =   1485
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Revision Recepcion"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2040
      TabIndex        =   9
      Top             =   3240
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmAuditoriaCalidadRecepcionAvios.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmAuditoriaCalidadRecepcionAvios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flg_Status_aprobacion As String
Dim strSQL As String

Public codigo As String
Public descripcion As String

Public Cod_Almacen As String
Public Num_MovStk As String
Public Num_Secuencia As String
Public Cantidad As Long




Private Sub Form_Load()
optAprobado.Value = True
Me.TxtCod_Motivo.Enabled = False
Me.TxtDes_Motivo.Enabled = False
Me.TxtCantidadDesaprobada.Enabled = False
Me.TxtAql.Enabled = False
'Me.TxtPorcDesaprobacion.Enabled = False
'Me.TxtObservaciones.Enabled = False

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "ACEPTAR"
        SALVAR_DATOS
        Unload Me
    Case "CANCELAR"
        Unload Me
End Select
End Sub

Private Sub optAprobado_Click()
    flg_Status_aprobacion = "S"
    Me.TxtCod_Motivo.Enabled = False
    Me.TxtDes_Motivo.Enabled = False
    Me.TxtCantidadDesaprobada.Enabled = False
    Me.TxtAql.Enabled = False
    'Me.TxtPorcDesaprobacion.Enabled = False
    'Me.TxtObservaciones.Enabled = False
    Limpia_Desaprobado
    
End Sub

Private Sub optDesaprobado_Click()
    flg_Status_aprobacion = "N"
    Me.TxtCod_Motivo.Enabled = True
    Me.TxtDes_Motivo.Enabled = True
    Me.TxtCantidadDesaprobada.Enabled = True
    Me.TxtAql.Enabled = True
    'Me.TxtPorcDesaprobacion.Enabled = True
    Me.TxtObservaciones.Enabled = True
    Limpia_Desaprobado
    Me.TxtCod_Motivo.SetFocus
End Sub
Private Sub Limpia_Desaprobado()
    Me.TxtCod_Motivo.Text = ""
    Me.TxtDes_Motivo.Text = ""
    Me.TxtCantidadDesaprobada.Text = ""
    Me.TxtAql.Text = ""
    'Me.TxtPorcDesaprobacion.Text = ""
    Me.TxtObservaciones.Text = ""
End Sub


Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    Dim Rs As New ADODB.Recordset
    On Error GoTo Salvar_DatosErr
    Dim strSQL As String
    
    Con.ConnectionString = cConnect
    Con.Open
    
        Con.BeginTrans
            'IIf(TxtPorcDesaprobacion = "", "0", TxtPorcDesaprobacion) & "','" & _

        strSQL = "EXEC UP_ACTUALIZA_AUDITORIACALIDADRECEPCIONAVIOS '" & _
        Cod_Almacen & "','" & _
        Num_MovStk & "','" & _
        Num_Secuencia & "','" & _
        flg_Status_aprobacion & "','" & _
        DTPicker1.Value & "','" & _
        TxtCod_Motivo & "','" & _
        IIf(TxtCantidadDesaprobada = "", "0", TxtCantidadDesaprobada) & "','" & _
        IIf(TxtAql = "", "0", TxtAql) & "','" & _
        TxtObservaciones & "','" & _
        vusu & "'"

        Con.Execute strSQL
       
        Con.CommitTrans
        MsgBox "Los datos fueron procesados con Èxito.", vbInformation, "Mensaje"
        

    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler err, "Salvar_Datos"
End Sub


Private Sub TxtAql_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtObservaciones.SetFocus
End If
End Sub

Private Sub TxtCantidadDesaprobada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        If Me.TxtCantidadDesaprobada <> "" Then
            If CLng(Me.TxtCantidadDesaprobada) >= 1 And Cantidad >= 1 Then
                If CLng(Me.TxtCantidadDesaprobada) > Cantidad Then
                    MsgBox "Los Cantidad Desaprobada debe ser menor o igual a :" & Cantidad, vbInformation, "Mensaje"
                    Exit Sub
                End If
                Me.TxtPorcDesaprobacion = Round((CLng(Me.TxtCantidadDesaprobada) / Cantidad) * 100, 2)
            ElseIf CLng(Me.TxtCantidadDesaprobada) = 0 Then
                    MsgBox "Los Cantidad Desaprobada debe ser mayor a:  0 y menor o igual a: " & Cantidad, vbInformation, "Mensaje"
                    Exit Sub
            End If
        End If
        TxtAql.SetFocus
    End If
End Sub

Private Sub TxtCod_Motivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.TxtCod_Motivo.Text) = "" Then
            Call Me.BuscaMotivo(3)
        Else
            Call Me.BuscaMotivo(1)
        End If
    End If

End Sub

Public Sub BuscaMotivo(opcion As Integer)
Dim sField As String, iRows As Long
Dim rstAux As ADODB.Recordset
    
    strSQL = " SELECT RTRIM(Cod_Desaprobacion_CC_Proveedor) as 'Codigo' , RTRIM(Descripcion) AS 'DescripciÛn' FROM CC_Motivos_Desaprobacion_Recepcion_Proveedor WHERE "
    TxtCod_Motivo = Trim(TxtCod_Motivo)
    TxtDes_Motivo = Trim(TxtDes_Motivo)
    'sField = TxtCod_Motivo
    Select Case opcion
    Case 1: strSQL = strSQL & " Cod_Desaprobacion_CC_Proveedor   like '%" & Trim(TxtCod_Motivo.Text) & "%'  "
    Case 2: strSQL = strSQL & " Descripcion  like '%" & Trim(TxtDes_Motivo.Text) & "%' "
    Case 3: strSQL = " SELECT RTRIM(Cod_Desaprobacion_CC_Proveedor) as 'Codigo' , RTRIM(Descripcion) AS 'DescripciÛn' FROM CC_Motivos_Desaprobacion_Recepcion_Proveedor "
    End Select
    
    TxtCod_Motivo = ""
    TxtDes_Motivo = ""
    With frmBusqGeneral
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        
        codigo = ""
        descripcion = ""
        
        iRows = .gexList.RowCount
        Set rstAux = .gexList.ADORecordset
        If .gexList.RowCount > 1 Then
            .Show vbModal
        ElseIf .gexList.RowCount = 1 Then
            codigo = .gexList.Value(.gexList.Columns("CODIGO").Index)
            descripcion = .gexList.Value(.gexList.Columns("DESCRIPCIÛN").Index)
        End If
        
        If codigo <> "" Then
            TxtCod_Motivo = RTrim(codigo)
            TxtDes_Motivo = RTrim(descripcion)
            TxtCantidadDesaprobada.SetFocus
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub

Private Sub TxtObservaciones_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FunctButt1.SetFocus
End If
End Sub

'Private Sub TxtPorcDesaprobacion_KeyPress(KeyAscii As Integer)
' If KeyAscii = 13 Then
'    TxtObservaciones.SetFocus
' End If
'End Sub
