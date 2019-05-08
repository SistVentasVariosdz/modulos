VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRepEstCliTem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte"
   ClientHeight    =   2505
   ClientLeft      =   4575
   ClientTop       =   3585
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frFecha 
      Height          =   2295
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   4815
      Begin VB.TextBox Txtcod_Version 
         Height          =   330
         Left            =   840
         TabIndex        =   13
         Top             =   180
         Width           =   350
      End
      Begin VB.TextBox TxtDes_Version 
         Height          =   330
         Left            =   1200
         TabIndex        =   12
         Top             =   180
         Width           =   3465
      End
      Begin VB.Frame Frame3 
         Caption         =   "Fecha que Envio el Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   4575
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   285
            Left            =   840
            TabIndex        =   10
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   503
            _Version        =   393216
            Format          =   62324737
            CurrentDate     =   38458
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   495
         Left            =   2415
         TabIndex        =   8
         Top             =   1530
         Width           =   1275
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Imprimir"
         Height          =   495
         Left            =   1020
         TabIndex        =   7
         Top             =   1530
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Version:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   280
         Width           =   615
      End
   End
   Begin VB.Frame frPrincipal 
      Height          =   2475
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4125
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Imprimir"
         Height          =   495
         Left            =   660
         TabIndex        =   5
         Top             =   1890
         Width           =   1305
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   495
         Left            =   2055
         TabIndex        =   4
         Top             =   1905
         Width           =   1275
      End
      Begin VB.Frame Frame2 
         Caption         =   "Seleccione Reporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3855
         Begin VB.OptionButton OptEstiloSeleccionado 
            Caption         =   "Estilo Seleccionado / Ult.Iteración"
            Height          =   450
            Left            =   615
            TabIndex        =   11
            Top             =   880
            Width           =   2820
         End
         Begin VB.OptionButton optResCom 
            Caption         =   "Resumido Comercial"
            Height          =   315
            Left            =   615
            TabIndex        =   3
            Top             =   270
            Width           =   1890
         End
         Begin VB.OptionButton optDetDes 
            Caption         =   "Detallado Desarrollo"
            Height          =   450
            Left            =   615
            TabIndex        =   2
            Top             =   525
            Value           =   -1  'True
            Width           =   1860
         End
      End
   End
End
Attribute VB_Name = "frmRepEstCliTem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public varCod_TemCli As String
Public varCod_EstPro As String
Public varAbr_Cliente As String
Public varDes_Cliente As String
Public varNom_TemCli As String
Public varNumCot As Integer
Public varObs As String
Public varCod_EstCli As String

Public Codigo As String, Descripcion As String, TipoAdd As String

Dim StrSQL As String

Private Sub cmdAceptar_Click()

  'If OptEstiloSeleccionado Then
    'If MsgBox("Desea Incorporar al Seguimiento Wip Protos", vbInformation + vbYesNo, "ADVERTENCIA") = vbYes Then
      'frPrincipal.Visible = False
      'frFecha.Visible = True
      'frFecha.Top = 0
      'Txtcod_Version.SetFocus
      'Exit Sub
    'End If
  'End If
  
  Call GeneraReportes
End Sub

Private Sub cmdCancel_Click()
  frFecha.Visible = False
  frPrincipal.Visible = True
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Sub GeneraReportes()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String
Dim StrSQL As String
Dim vCod_Version As String
Dim vCod_Cliente As String
Dim vNum_Iteracion As Integer

    If Me.optResCom.Value = True Then
        Ruta = vRuta & "\prototipo.xlt"
    ElseIf Me.OptEstiloSeleccionado Then
        Ruta = vRuta & "\PROTOTIPOD_Estilo.xlt"
    Else
        Ruta = vRuta & "\prototipoD.xlt"
    End If
    
    StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(varAbr_Cliente) & "'"
    vCod_Cliente = CStr(DevuelveCampo(StrSQL, cCONNECT))
    
    vCod_Version = DevuelveCampo("select cod_version_costeo from tg_estcliest where cod_cliente ='" & vCod_Cliente & "' and cod_temcli='" & varCod_TemCli & "' and cod_estcli='" & varCod_EstCli & "' and cod_estpro='" & varCod_EstPro & "'", cCONNECT)
    vNum_Iteracion = DevuelveCampo("select num_iteracion from tg_estcliest where cod_cliente ='" & vCod_Cliente & "' and cod_temcli='" & varCod_TemCli & "' and cod_estcli='" & varCod_EstCli & "' and cod_estpro='" & varCod_EstPro & "'", cCONNECT)
    
    Set oo = CreateObject("excel.application")
    oo.workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.run "Reporte", vCod_Cliente, varCod_TemCli, varNumCot, cCONNECT, vemp, varDes_Cliente, varNom_TemCli, varObs, vusu, varCod_EstCli, varCod_EstPro, vCod_Version, vNum_Iteracion
    Set oo = Nothing
Exit Sub
Resume
hand:
    ErrorHandler Err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub cmdPrint_Click()
'Dim Cod_Version As String
Dim varCod_Cliente As String

On Error GoTo hand

  Call GeneraReportes
  
  'Cod_Version = DevuelveCampo("Select  b.Cod_Version_Costeo from tg_estclitem a ,  tg_estcliest b where a.num_solicitud_cons = " & varNumCot & "  and a.cod_cliente = b.cod_cliente and  a.cod_temcli = b.cod_temcli and a.cod_estcli = b.cod_estcli  and   b.cod_estpro = '" & varCod_EstPro & "'", cCONNECT)
  'Call ExecuteCommandSQL(cCONNECT, "es_incorpora_estilo_version_al_wip '" & varCod_EstPro & "','" & Cod_Version & "','" & vusu & "','" & dtpFecha & "'")
  varCod_Cliente = DevuelveCampo("select cod_cliente from tg_cliente where abr_cliente ='" & varAbr_Cliente & "'", cCONNECT)
  Call ExecuteCommandSQL(cCONNECT, "es_incorpora_estilo_version_al_wip '" & varCod_EstPro & "','" & TxtCod_Version.Text & "','" & vusu & "','" & dtpFecha & "','" & varCod_Cliente & "','" & Me.varCod_TemCli & "','" & Me.varCod_EstCli & "'")
  
  Call cmdCancel_Click
Exit Sub

hand:
    ErrorHandler Err, "GeneraReportes"
End Sub

Private Sub Form_Load()
  dtpFecha = Date
End Sub

Private Sub Txtcod_Version_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(TxtCod_Version.Text) = "" Then
            Call BUSCA_VERSION
        Else
            StrSQL = "SELECT Des_Version FROM ES_ESTPROVER WHERE Cod_EstPro = '" & varCod_EstPro & "' AND Cod_Version='" & Trim(TxtCod_Version.Text) & "'"
            txtDes_Version.Text = DevuelveCampo(StrSQL, cCONNECT)
        End If
    End If
End Sub

Private Sub TxtDes_Version_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtDes_Version.Text)) < 5 Then
            Call MsgBox("La descripción debe tener como mínimo 5 caracteres. Sirvase verificar", vbInformation)
            Exit Sub
        Else
            StrSQL = "SELECT Cod_Version FROM ES_ESTPROVER WHERE Cod_EstPro = '" & varCod_EstPro & "' AND Des_Version LIKE '" & Trim(txtDes_Version.Text) & "%'"
            TxtCod_Version.Text = DevuelveCampo(StrSQL, cCONNECT)
        End If
    End If
End Sub

Public Sub BUSCA_VERSION()
    Dim oTipo As New frmBusqGeneral3
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT Cod_Version AS Código, Des_Version as Descripcion FROM ES_ESTPROVER WHERE Cod_EstPro = '" & varCod_EstPro & "'"
    oTipo.Cargar_Datos
    oTipo.DGridLista.Columns(1).Width = 850
    oTipo.DGridLista.Columns(2).Width = 4000
    
    oTipo.Show 1
    If Codigo <> "" Then
        TxtCod_Version.Text = Codigo
        txtDes_Version.Text = Descripcion
        dtpFecha.SetFocus
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub
