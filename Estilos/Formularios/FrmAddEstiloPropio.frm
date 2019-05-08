VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmAddEstiloPropio 
   Caption         =   "Asignar Estilo Propio al Estilo Cliente"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2880
      TabIndex        =   9
      Top             =   3960
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmAddEstiloPropio.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   14
      Top             =   1440
      Width           =   8295
      Begin VB.OptionButton OptOtroEstCli 
         Caption         =   "Otro Estilo Cliente"
         Height          =   255
         Left            =   6600
         TabIndex        =   35
         Top             =   300
         Width           =   1575
      End
      Begin VB.OptionButton OptExiste 
         Caption         =   "Estilo Propio existente"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   300
         Width           =   1935
      End
      Begin VB.OptionButton OptNuevo 
         Caption         =   "Nuevo Estilo Propio"
         Height          =   255
         Left            =   4800
         TabIndex        =   17
         Top             =   300
         Width           =   1695
      End
      Begin VB.OptionButton OptOtraTempo 
         Caption         =   "Estilos Clientes otras Temporadas"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Value           =   -1  'True
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1425
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8295
      Begin VB.TextBox TxtDes_TemCli 
         Height          =   285
         Left            =   2520
         TabIndex        =   30
         Top             =   600
         Width           =   5175
      End
      Begin VB.TextBox TxtCod_TemCli 
         Height          =   285
         Left            =   1680
         TabIndex        =   29
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TxtAbr_Cliente 
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtNom_Cliente 
         Height          =   285
         Left            =   2520
         TabIndex        =   20
         Top             =   240
         Width           =   5175
      End
      Begin VB.TextBox TxtDes_EstCli 
         Height          =   285
         Left            =   3720
         TabIndex        =   13
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox TxtCod_EstCli 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Temporada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   31
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   22
         Top             =   330
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estilo Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   1125
      End
   End
   Begin VB.Frame FraExistente 
      Caption         =   "Estilo Propio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   30
      TabIndex        =   18
      Top             =   150
      Width           =   8295
      Begin VB.TextBox TxtCod_EstCliOtro 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtNum_Veces 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Text            =   "1"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Txtcod_EstPro 
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TxtDes_EstPro 
         Height          =   285
         Left            =   2520
         TabIndex        =   2
         Top             =   735
         Width           =   4815
      End
      Begin VB.Label LblEstCli 
         AutoSize        =   -1  'True
         Caption         =   "Estilo Cliente"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   34
         Top             =   465
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Veces :"
         Height          =   195
         Left            =   360
         TabIndex        =   28
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estilo Propio"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   19
         Top             =   840
         Width           =   870
      End
   End
   Begin VB.Frame FraNuevo 
      Caption         =   "Nuevo Estilo Propio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   23
      Top             =   2220
      Visible         =   0   'False
      Width           =   8295
      Begin VB.TextBox TxtDes_UsuPre 
         Height          =   285
         Left            =   2040
         TabIndex        =   33
         Top             =   1320
         Width           =   5295
      End
      Begin VB.TextBox TxtCod_UsuPre 
         Height          =   285
         Left            =   1200
         TabIndex        =   32
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TxtDes_TipPre 
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Top             =   600
         Width           =   5295
      End
      Begin VB.TextBox TxtCod_TipPre 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TxtCod_GruTalla 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TxtDes_GruTalla 
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Top             =   960
         Width           =   5295
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Usu. Pren :"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1395
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Grupo Tallas"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   1050
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   330
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Prenda"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   24
         Top             =   700
         Width           =   870
      End
   End
End
Attribute VB_Name = "FrmAddEstiloPropio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vCod_Cliente As String, vCod_TemCli As String, vCod_estCli As String, sDes_Estilo As String, vDes_Tela As String
Dim strSQL As String
Dim vOpcion As String
Public Codigo As String, Descripcion As String, TipoAdd As String, tipoAdd2 As String

Dim varCod_EstPro As String

Private Sub Form_Load()
vOpcion = "T"
'LlenaCombo Me.cmbTipPre, "Select Des_TipPre  + space(100) +Cod_TipPre from tg_tippre order by Des_TipPre  ", cCONNECT
'LlenaCombo Me.cboCod_UsuPre, "Select  Des_UsuPre + space(100)+ Cod_UsuPre  from TG_USUPRENDAS", cCONNECT
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call Graba_EP
Case "CANCELAR"
    Unload Me
End Select
End Sub

Sub Graba_EP()
Dim vMensaje As Variant
On Error GoTo hand

If vOpcion <> "N" Then
    strSQL = "EXEC UP_MAN_ESTCLIEST 'I','" & _
            Me.vCod_Cliente & "','" & _
            Me.vCod_TemCli & "','" & _
            Me.vCod_estCli & "','" & _
            Txtcod_EstPro.Text & "'," & _
            txtNum_Veces.Text

    ExecuteCommandSQL cCONNECT, strSQL
    
    varCod_EstPro = Txtcod_EstPro.Text
Else
    
    varCod_EstPro = DevuelveCampo("EXEC UP_Es_EsPro 'i','','','','','" & Me.TxtDescripcion & "','" & TxtCod_TipPre.Text & "','" & TxtCod_GruTalla & "','','','" & TxtCod_UsuPre.Text & "'", cCONNECT)
    If varCod_EstPro <> "" Then
        strSQL = "select Des_estpro from es_estpro where Cod_EstPro='" & varCod_EstPro & "' and Cod_tippre='" & Trim(TxtCod_TipPre.Text) & "' and Cod_GruTal='" & TxtCod_GruTalla & "'"
        Call MsgBox("Se creo el estilo " & Trim(DevuelveCampo(strSQL, cCONNECT)) & " con el siguiente numero :" & varCod_EstPro, vbInformation)
    End If
    
    strSQL = "EXEC UP_MAN_ESTCLIEST 'I','" & _
            Me.vCod_Cliente & "','" & _
            Me.vCod_TemCli & "','" & _
            Me.vCod_estCli & "','" & _
            varCod_EstPro & "',1"

    ExecuteCommandSQL cCONNECT, strSQL
End If

If Val(DevuelveCampo("select count(*) from es_estprover where cod_estpro='" & varCod_EstPro & "'", cCONNECT)) = 0 Then
    Call Asigna_Version
Else
    vMensaje = MsgBox("¿Desea Crear Nueva Version?", vbYesNo)
    If vMensaje = vbNo Then
        Load FrmShowVersiones
        Set FrmShowVersiones.oParent = Me
        FrmShowVersiones.vCod_Cliente = Me.vCod_Cliente
        FrmShowVersiones.vCod_TemCli = Me.vCod_TemCli
        FrmShowVersiones.vCod_estCli = Me.vCod_estCli
        FrmShowVersiones.vCod_estPro = varCod_EstPro
        FrmShowVersiones.CARGA_GRID
        FrmShowVersiones.Show vbModal
        Set FrmShowVersiones = Nothing
    Else
        Call Asigna_Version
    End If
End If

Unload Me
Exit Sub
hand:
    Unload Me
    MsgBox Err.Description, vbCritical, "Grabar EP"
End Sub

Private Sub OptExiste_Click()
vOpcion = "E"
FraExistente.Visible = True
FraNuevo.Visible = False
Call Limpia_Text
TxtCod_EstCliOtro.Visible = False
LblEstCli(3).Visible = False
Txtcod_EstPro.SetFocus
End Sub

Private Sub OptNuevo_Click()
vOpcion = "N"
FraExistente.Visible = False
FraNuevo.Visible = True
Call Limpia_Text
TxtDescripcion.SetFocus
End Sub

Private Sub OptOtraTempo_Click()
vOpcion = "T"
FraExistente.Visible = True
FraNuevo.Visible = False
Call Limpia_Text
TxtCod_EstCliOtro.Visible = False
LblEstCli(3).Visible = False
Txtcod_EstPro.SetFocus
End Sub

Private Sub OptOtroEstCli_Click()
vOpcion = "O"
FraExistente.Visible = True
FraNuevo.Visible = False
Call Limpia_Text
TxtCod_EstCliOtro.Visible = True
LblEstCli(3).Visible = True
TxtCod_EstCliOtro.SetFocus
End Sub

Private Sub TxtCod_EstCliOtro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtCod_EstPro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If vOpcion = "E" Then
        If Trim(Txtcod_EstPro.Text) = "" Then
            'cmdBusCliente_Click
        Else
            Txtcod_EstPro.Text = Right("00000" & Trim(Txtcod_EstPro.Text), 5)
            strSQL = "SELECT  Des_EstPro FROM ES_EstPro WHERE Cod_EstPro='" & Txtcod_EstPro.Text & "'"
            TxtDes_EstPro.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
            txtNum_Veces.SetFocus
        End If
    ElseIf vOpcion = "T" Then
        Call Busca_Estilos(vCod_estCli)
    ElseIf vOpcion = "O" Then
        Call Busca_Estilos(TxtCod_EstCliOtro)
    End If
End If
End Sub

Private Sub TxtCod_GruTalla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_GruTal(1)
End If
End Sub

Private Sub TxtCod_TipPre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_TipoPrenda(1)
End If
End Sub

Private Sub TxtCod_UsuPre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_UsuPrenda(1)
End If
End Sub

Private Sub txtDes_estpro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If vOpcion = "E" Then
        If Len(TxtDes_EstPro) > 4 Then
            strSQL = "SELECT Cod_EstPro FROM ES_EstPro WHERE Des_EstPro LIKE '" & Trim(TxtDes_EstPro.Text) & "%'"
            Txtcod_EstPro.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
            strSQL = "SELECT  Des_EstPro FROM ES_EstPro WHERE Cod_EstPro='" & Txtcod_EstPro.Text & "'"
            TxtDes_EstPro.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
            'txtCod_TemCli.SetFocus
        Else
            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
            TxtDes_EstPro.SetFocus
        End If
    ElseIf vOpcion = "T" Then
        Call Busca_Estilos(vCod_estCli)
    End If
End If
End Sub

Sub Busca_Estilos(vEstCli As String)
Dim rstAux As ADODB.Recordset
On Error GoTo Fin
Dim iCol As Long
    
    strSQL = "exec es_muestra_estilos_clientes_propios_otras_temporadas '" & vCod_Cliente & "','" & vEstCli & "'"
    
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        .Caption = "Seleccionar Estilo Propio"
        Codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Estilo_Propio").Width = 900
        .DGridLista.Columns("Nombre").Width = 5000
        
        If rstAux.RecordCount = 1 Then
            Codigo = Trim(rstAux!Estilo_Propio)
            Descripcion = Trim(rstAux!Nombre)
            'MsgBox "No existen Est.Propios en otras Temporadas", vbCritical
        ElseIf rstAux.RecordCount > 1 Then
            .Show vbModal
        End If
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            Txtcod_EstPro = Codigo
            TxtDes_EstPro = Descripcion
            txtNum_Veces.SetFocus
        End If
    End With
    Codigo = "": Descripcion = ""
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Busca Estilos"
End Sub

Sub Limpia_Text()
Txtcod_EstPro.Text = ""
TxtDes_EstPro.Text = ""
txtNum_Veces.Text = "1"
TxtDescripcion.Text = Me.sDes_Estilo
TxtCod_TipPre.Text = ""
TxtDes_TipPre.Text = ""
TxtCod_UsuPre.Text = ""
TxtDes_UsuPre.Text = ""
TxtCod_GruTalla.Text = ""
TxtDes_GruTalla.Text = ""
End Sub

Private Sub TxtDes_GruTalla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call Busca_GruTal(2)
End If
End Sub

Private Sub TxtDes_TipPre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_TipoPrenda(2)
End If
End Sub

Private Sub TxtDes_UsuPre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_UsuPrenda(2)
End If
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtDescripcion.Text = Trim(TxtDescripcion.Text)
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtNum_Veces_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FunctButt1.SetFocus
Else
    Call SoloNumeros(txtNum_Veces, KeyAscii, False, 0, 3)
End If
End Sub

Private Sub txtNum_Veces_LostFocus()
If Trim(txtNum_Veces.Text) = "" Then
    txtNum_Veces.Text = 1
End If
End Sub

Sub Busca_GruTal(opcion As Integer)
Dim vMensaje As Variant
Dim rstAux As ADODB.Recordset
On Error GoTo Fin
Dim iCol As Long
    
    If opcion = 1 Then
        strSQL = "Select  cod_grutal as Codigo, des_grutal as Descripcion from es_tallas where cod_grutal like '%" & Trim(TxtCod_GruTalla.Text) & "%' order by cod_grutal"
    Else
        strSQL = "Sm_Muestra_Ayuda_GRupo_Tallas '" & Trim(TxtDes_GruTalla.Text) & "'"
    End If
    
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        .Caption = "Seleccionar Grupo Talla"
        'Codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Codigo").Width = 900
        .DGridLista.Columns("Descripcion").Width = 5000
        
        If rstAux.RecordCount = 1 Then
            Codigo = rstAux!Codigo
            Descripcion = rstAux!Descripcion
        End If
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            TxtCod_GruTalla = Codigo
            TxtDes_GruTalla = Descripcion
            TxtCod_UsuPre.SetFocus
        End If
        If opcion = 2 And Codigo = "" Then
            vMensaje = MsgBox("No se pudo encontrar la cadena, ¿Desea crear nuevo grupo Tallas?", vbYesNo)
            If vMensaje = vbNo Then Exit Sub
            Call Graba_Grupo
        End If
    End With
    Codigo = "": Descripcion = ""
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Busca Grupo Talla"
End Sub

Sub Asigna_Version()
Load FrmAddVersion
FrmAddVersion.vCodEstPro = varCod_EstPro
FrmAddVersion.sCod_EstCli = Me.vCod_estCli
FrmAddVersion.sCod_Cliente = Me.vCod_Cliente
FrmAddVersion.vCodTemporada = Me.vCod_TemCli
FrmAddVersion.vDes_Tela = Me.vDes_Tela
FrmAddVersion.TxtDes_TelaCliente = Me.vDes_Tela

If DevuelveCampo("select count(*) from es_estprover where cod_estpro='" & varCod_EstPro & "'", cCONNECT) = 0 Then
    FrmAddVersion.TxtCodigo = DevuelveCampo("Select cod_version_default from tg_control", cCONNECT)
    FrmAddVersion.TxtDescripcion = DevuelveCampo("Select des_version_default from tg_control", cCONNECT)
End If
Call BuscaCombo(DevuelveCampo("select Tip_Version_Default from tg_control", cCONNECT), 1, FrmAddVersion.dbcTipo)
FrmAddVersion.Tipo_Busqueda = "3"
FrmAddVersion.TipoBusq = "3"
FrmAddVersion.Show vbModal
Set FrmAddVersion = Nothing
End Sub

Sub Busca_TipoPrenda(opcion As Integer)
Dim rstAux As ADODB.Recordset
On Error GoTo Fin
Dim iCol As Long
    
    strSQL = "Select Cod_TipPre as Codigo, Des_TipPre as Descripcion from tg_tippre"
    
    If opcion = 1 Then
        strSQL = strSQL & " where Cod_TipPre like '%" & Trim(TxtCod_TipPre.Text) & "%' order by Cod_TipPre"
    Else
        strSQL = strSQL & " where Des_TipPre like '%" & Trim(TxtDes_TipPre.Text) & "%' order by Des_TipPre"
    End If
    
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        .Caption = "Seleccionar Tipo Prenda"
        Codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Codigo").Width = 900
        .DGridLista.Columns("Descripcion").Width = 5000
        
        If rstAux.RecordCount = 1 Then
            Codigo = rstAux!Codigo
            Descripcion = rstAux!Descripcion
        End If
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            TxtCod_TipPre = Codigo
            TxtDes_TipPre = Descripcion
            TxtCod_GruTalla.SetFocus
        End If
    End With
    Codigo = "": Descripcion = ""
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Busca Tipo Prenda"
End Sub

Sub Busca_UsuPrenda(opcion As Integer)
Dim rstAux As ADODB.Recordset
On Error GoTo Fin
Dim iCol As Long
    
    strSQL = "Select Cod_UsuPre as Codigo, Des_UsuPre as Descripcion from TG_USUPRENDAS"
    
    If opcion = 1 Then
        strSQL = strSQL & " where Cod_UsuPre like '%" & Trim(TxtCod_UsuPre.Text) & "%' order by Cod_UsuPre"
    Else
        strSQL = strSQL & " where Des_UsuPre like '%" & Trim(TxtDes_UsuPre.Text) & "%' order by Des_UsuPre"
    End If
    
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        .Caption = "Seleccionar Usuario Prenda"
        Codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Codigo").Width = 900
        .DGridLista.Columns("Descripcion").Width = 5000
        
        If rstAux.RecordCount = 1 Then
            Codigo = rstAux!Codigo
            Descripcion = rstAux!Descripcion
        End If
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            TxtCod_UsuPre = Codigo
            TxtDes_UsuPre = Descripcion
            FunctButt1.SetFocus
        End If
    End With
    Codigo = "": Descripcion = ""
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Busca Tipo Prenda"
End Sub

Sub Graba_Grupo()
Dim Rs As New ADODB.Recordset
On Error GoTo errGrabar

Set Rs = Nothing
Rs.CursorLocation = adUseClient

strSQL = "ES_MANTE_GRUPO_TALLAS '" & Trim(TxtDes_GruTalla) & "'"

Rs.Open strSQL, cCONNECT
If Rs.RecordCount Then
    TxtCod_GruTalla = Trim(Rs!Codigo)
    TxtDes_GruTalla = Trim(Rs!Descripcion)
    TxtCod_UsuPre.SetFocus
End If

Rs.Close
Exit Sub
errGrabar:
    Set Rs = Nothing
    MsgBox Err.Description, vbCritical
End Sub
