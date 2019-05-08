VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMan_RegProyVtaServText 
   Caption         =   "Mantenimiento de Proyeccion Ventas-Servicios Textiles"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.TextBox txtobservacion 
         Height          =   315
         Left            =   1950
         TabIndex        =   20
         Top             =   3120
         Width           =   7530
      End
      Begin VB.Frame Frame3 
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Visible         =   0   'False
         Width           =   9375
         Begin VB.CommandButton cmdBusTela 
            Caption         =   "..."
            Height          =   330
            Left            =   2865
            TabIndex        =   19
            Tag             =   "..."
            Top             =   120
            Width           =   360
         End
         Begin VB.TextBox txtdes_tela 
            Height          =   285
            Left            =   3330
            TabIndex        =   18
            Top             =   150
            Width           =   4725
         End
         Begin VB.TextBox txtcod_tela 
            Height          =   285
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   17
            Top             =   150
            Width           =   1005
         End
         Begin VB.Label Label5 
            Caption         =   "Tela:"
            Height          =   210
            Left            =   120
            TabIndex        =   16
            Top             =   200
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   9375
         Begin VB.TextBox txtOPCION_COD 
            Height          =   285
            Left            =   1770
            TabIndex        =   13
            Top             =   150
            Width           =   1245
         End
         Begin VB.TextBox txtHILADO_DES 
            Height          =   285
            Left            =   3030
            TabIndex        =   12
            Top             =   150
            Width           =   4995
         End
         Begin VB.Label Label4 
            Caption         =   "Codigo  Hilado:"
            Height          =   210
            Left            =   120
            TabIndex        =   14
            Top             =   200
            Width           =   1575
         End
      End
      Begin VB.TextBox txtkilos 
         Height          =   315
         Left            =   1950
         TabIndex        =   9
         Top             =   1580
         Width           =   1530
      End
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   315
         Left            =   1950
         TabIndex        =   5
         Top             =   750
         Width           =   690
      End
      Begin VB.TextBox txtNom_Cliente 
         Height          =   315
         Left            =   2700
         TabIndex        =   4
         Top             =   750
         Width           =   3840
      End
      Begin VB.TextBox TxtDes_TipoVenta 
         Height          =   315
         Left            =   2700
         TabIndex        =   2
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox TxtCod_TipoVenta 
         Height          =   315
         Left            =   1950
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPInicio 
         Height          =   300
         Left            =   1950
         TabIndex        =   7
         Top             =   1185
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         Format          =   94109697
         CurrentDate     =   37988
      End
      Begin VB.Label Label6 
         Caption         =   "Observacion:"
         Height          =   210
         Left            =   240
         TabIndex        =   21
         Top             =   3195
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Kilos Requeridos:"
         Height          =   210
         Left            =   240
         TabIndex        =   10
         Top             =   1650
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Requerimiento: "
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   1230
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         Height          =   210
         Left            =   240
         TabIndex        =   6
         Top             =   810
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Tipo de Venta :"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   375
         Width           =   1125
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3787
      TabIndex        =   22
      Top             =   3720
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmMan_RegProyVtaServText.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmMan_RegProyVtaServText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public codigo As String, sOpcion As String, Sid_proyeccion As Integer
Public Descripcion As String, TipoAdd As String
Dim strSQL As String
Dim SGRUPO As String

Private Sub Form_Load()
dtpInicio.Value = Date + 1
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim sMessage As Long
Dim strSQL As String

On Error GoTo dprError

Select Case ActionName
  Case "ACEPTAR"
     
    If TxtCod_TipoVenta.Text = "" Then
        MsgBox "Debe ingresar el Tipo de Venta"
        Exit Sub
    End If
    
    If txtkilos.Text = "" Then
        MsgBox "Debe ingresar los Kilos Requeridos"
        Exit Sub
    End If
    
  
    If MsgBox("Esta seguro de grabar... ", vbYesNo, "IMPORTANTE") = vbYes Then
        Salvar_Datos
        Unload Me
      End If
    
    
  Case "CANCELAR"
      Unload Me

End Select

Exit Sub


dprError:

errores err.Number
End Sub

Sub Salvar_Datos()
On Error GoTo ErrSalvarDatos
Dim vCod_Cliente_Tex As String


    
    strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(txtAbr_Cliente.Text) & "'"
    vCod_Cliente_Tex = DevuelveCampo(strSQL, cCONNECT)

    strSQL = "exec ventas_up_act_proyeccion_textil_status '" & sOpcion & "'," & Sid_proyeccion & ",'" & TxtCod_TipoVenta.Text & "','" & vCod_Cliente_Tex & "','" & dtpInicio.Value & "','" & txtkilos.Text & "','" & txtOPCION_COD.Text & "','" & txtcod_tela.Text & "','" & TxtObservacion.Text & "','" & vusu & "'"
    ExecuteSQL cCONNECT, strSQL
    
        
Exit Sub
ErrSalvarDatos:
    ErrorHandler err, "SALVAR_DATOS"
End Sub

Private Sub TxtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtAbr_Cliente.Text) = "" Then
            Call BUSCA_CLIENTE(3)
        Else
            Call BUSCA_CLIENTE(1)
        End If
    End If

End Sub



Public Sub BUSCA_CLIENTE(Tipo As Integer)
    Select Case Tipo
        Case 1:
                    strSQL = "EXEC TI_BUSCA_CLIENTE 1,'" & Trim(Me.txtAbr_Cliente.Text) & "','','" & vusu & "'"
                    Me.txtNom_Cliente.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                    txtkilos.SetFocus
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim RS As Object
                    Set RS = CreateObject("ADODB.Recordset")
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 2,'','" & Trim(txtNom_Cliente.Text) & "','" & vusu & "'"
                    Else
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 3,'','','" & vusu & "'"
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.DGridLista.Columns(2).Width = 3500
                    oTipo.Show 1
                    If codigo <> "" Then
                         Me.txtAbr_Cliente.Text = Trim(codigo)
                         Me.txtNom_Cliente.Text = Trim(Descripcion)
'                         OptCliPend.SetFocus
                         codigo = "": Descripcion = ""
                        txtkilos.SetFocus
                    End If
                    Set oTipo = Nothing
                    Set RS = Nothing
    End Select
    
End Sub

Private Sub txtcod_tela_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtcod_tela.Text) = "" Then
            MsgBox ("Sirvase ingresar un codigo de Item")
        Else
            txtcod_tela.Text = CompletaCodigo(Trim(txtcod_tela.Text), 8, 2)
            
            'Esta consulta es para obtener el Codigo de Cliente
            strSQL = "SELECT Des_Tela FROM TX_TELA WHERE Cod_Tela ='" & txtcod_tela.Text & "'"
            txtdes_tela.Text = DevuelveCampo(strSQL, cCONNECT)
        End If
            TxtObservacion.SetFocus
    End If

End Sub

Public Function CompletaCodigo(CodOrigen As String, longcodfinal As Integer, PosfinalCod As Integer) As String
' CodOrigen     = Es el codigo que sera pasado por parametro
' LongCodFinal  = Es el tamaño del Codigo a devolver
' PosFinalCod   = Es la posicion de la 1era parte del codigo
    Dim Contador As Integer
    CompletaCodigo = Mid(CodOrigen, 1, PosfinalCod)
    For Contador = 1 To longcodfinal - Len(CodOrigen)
        CompletaCodigo = CompletaCodigo & "0"
    Next
    CompletaCodigo = CompletaCodigo & Right(CodOrigen, Len(CodOrigen) - PosfinalCod)
End Function

Private Sub TxtCod_TipoVenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(TxtCod_TipoVenta.Text) = "" Then
            Call BUSCA_TipoVenta(3)
        Else
            Call BUSCA_TipoVenta(1)
        End If
    End If

End Sub

Public Sub BUSCA_TipoVenta(Tipo As Integer)
    Select Case Tipo
        Case 1:
                    strSQL = "EXEC TI_BUSCA_TIPOS_VENTA 1,'" & Trim(Me.TxtCod_TipoVenta.Text) & "',''"
                    Me.TxtDes_TipoVenta.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral3
                    Dim RS As Object
                    Set RS = CreateObject("ADODB.Recordset")
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "EXEC TI_BUSCA_TIPOS_VENTA 2,'','" & Trim(TxtDes_TipoVenta.Text) & "'"
                    Else
                        oTipo.SQuery = "EXEC TI_BUSCA_TIPOS_VENTA 3,'',''"
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.gexLista.Columns(2).Width = 3500
                    oTipo.Show 1
                    If codigo <> "" Then
                         Me.TxtCod_TipoVenta.Text = Trim(codigo)
                         Me.TxtDes_TipoVenta.Text = Trim(Descripcion) 'TipoAdd
                         SGRUPO = Trim(TipoAdd)
'                         OptCliPend.SetFocus
                         codigo = "": Descripcion = "": TipoAdd = ""
                        If SGRUPO = "1" Then
                            Frame2.Visible = True
                            Frame3.Visible = False
                        Else
                            Frame3.Visible = True
                            Frame2.Visible = False
                        End If
                        txtAbr_Cliente.SetFocus
                    End If
                    Set oTipo = Nothing
                    Set RS = Nothing
    End Select
    
End Sub

Private Sub txtdes_tela_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtdes_tela.Text) = "" Then
             MsgBox ("Sirvase ingresar una Descripcion del Item")
        Else
            cmdBusTela_Click
        End If
'        Call Buscar
    End If

End Sub

Private Sub cmdBusTela_Click()
    Dim strSQL As String
    If Trim(txtcod_tela.Text) <> "" Then
        strSQL = "SELECT Cod_Tela as Código, Des_Tela as Descripción FROM TX_TELA WHERE Cod_Tela ='" & txtcod_tela.Text & "'"
    Else
        If Len(Trim(txtdes_tela.Text)) < 5 Then
            Call MsgBox("El Texto Ingresado debe contar con un mínimo de 5 caracteres", vbExclamation)
            Exit Sub
        Else
            strSQL = "SELECT Cod_Tela as Código, Des_Tela as Descripción FROM TX_TELA WHERE Des_Tela LIKE '" & Trim(txtdes_tela.Text) & "%'"
        End If
    End If
    
    Dim oTipo As New frmBusqGeneral
    Dim RS As Object
    Set RS = CreateObject("ADODB.Recordset")
    Set oTipo.oParent = Me
    oTipo.SQuery = strSQL
    oTipo.CARGAR_DATOS
    oTipo.Show 1
    If codigo <> "" Then
        txtcod_tela.Text = codigo
        txtdes_tela.Text = Descripcion
'        FunctBuscar.SetFocus
    End If
    Set oTipo = Nothing
    Set RS = Nothing
End Sub


Private Sub TxtDes_TipoVenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(TxtDes_TipoVenta.Text) = "" Then
            Call BUSCA_TipoVenta(3)
        Else
            Call BUSCA_TipoVenta(2)
        End If
    End If


End Sub

Private Sub txtHILADO_DES_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            Call BUSCAHILADO(2)
    End If

End Sub

Private Sub txtkilos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    If SGRUPO = "1" Then
        txtOPCION_COD.SetFocus
    Else
        txtcod_tela.SetFocus
    End If
    End If

End Sub

Private Sub TxtNom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtNom_Cliente.Text) = "" Then
            Call BUSCA_CLIENTE(3)
        Else
            Call BUSCA_CLIENTE(2)
        End If
    End If

End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       FunctButt1.SetFocus
    End If

End Sub

Private Sub txtOPCION_COD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            Call BUSCAHILADO(1)
    End If

End Sub


Private Sub BUSCAHILADO(opcion As Integer)
Dim sField As String
Dim iRows As Long
Dim rstAux As ADODB.Recordset
 
    strSQL = "SELECT COD_HILADO AS CODIGO, DESCRIPCION AS NOMBRE FROM HI_HILADOS WHERE "
    txtOPCION_COD = Trim(txtOPCION_COD)
    txtHILADO_DES = Trim(txtHILADO_DES)
    sField = txtOPCION_COD
    Select Case opcion
        Case 1: strSQL = strSQL & "COD_HILADO like '%" & txtOPCION_COD & "%'"
        Case 2: strSQL = strSQL & "DESCRIPCION like '%" & txtHILADO_DES & "%'"
    End Select
    
    txtOPCION_COD = Empty
    txtHILADO_DES = Empty
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
        .Caption = "Seleccionar - Hilado"
        .CARGAR_DATOS
        .DGridLista.Columns("Nombre").Width = 3500
        
        codigo = Empty
        Descripcion = Empty
        
        iRows = .DGridLista.RowCount
        Set rstAux = .DGridLista.ADORecordset
        If .DGridLista.RowCount > 1 Then
            .Show vbModal
        ElseIf .DGridLista.RowCount = 1 Then
            codigo = .DGridLista.Value(.DGridLista.Columns("CODIGO").Index)
            Descripcion = .DGridLista.Value(.DGridLista.Columns("NOMBRE").Index)
        End If
        
        If codigo <> "" Then
            txtOPCION_COD = RTrim(codigo)
            txtHILADO_DES = RTrim(Descripcion)
            TxtObservacion.SetFocus
        End If
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub
