VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form frmShowCierreTipoDiario_Ventas 
   Caption         =   "Cierre Mensual"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSubdiarios 
      Height          =   1485
      Left            =   60
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   5355
      Begin VB.TextBox txtDesSubdiario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1395
         TabIndex        =   11
         Top             =   180
         Width           =   3900
      End
      Begin VB.TextBox txtSubDiario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   915
         TabIndex        =   10
         Top             =   180
         Width           =   405
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   510
         Left            =   1440
         TabIndex        =   13
         Top             =   870
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmShowCierreTipoDiario_Ventas.frx":0000
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label lblSDStatus 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1410
         TabIndex        =   14
         Top             =   540
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "TipoDiario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   60
         TabIndex        =   12
         Top             =   195
         Width           =   735
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1500
      TabIndex        =   4
      Top             =   1140
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmShowCierreTipoDiario_Ventas.frx":0097
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.TextBox TxtAno 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   975
      TabIndex        =   2
      Top             =   435
      Width           =   615
   End
   Begin VB.TextBox TxtPeriodo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   450
      Width           =   405
   End
   Begin VB.TextBox TxtCod_TipoDiario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   975
      TabIndex        =   0
      Top             =   75
      Width           =   405
   End
   Begin VB.TextBox TxtDes_TipoDiario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1455
      TabIndex        =   1
      Top             =   75
      Width           =   3900
   End
   Begin VB.Label lblStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   480
      Width           =   2055
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   4560
      Top             =   1020
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Año"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   7
      Top             =   465
      Width           =   600
   End
   Begin VB.Label Label2 
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1740
      TabIndex        =   6
      Top             =   465
      Width           =   600
   End
   Begin VB.Label Label3 
      Caption         =   "TipoDiario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   5
      Top             =   90
      Width           =   735
   End
End
Attribute VB_Name = "frmShowCierreTipoDiario_Ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public codigo As String
Public descripcion As String
Public TipoAdd As String

Dim Strsql As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            CambiarEstado
        Case "CANCELAR"
            Unload Me
    End Select
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo ERRX
    Select Case ActionName
    Case "ACEPTAR"
        Strsql = "CN_CAMBIAR_STATUS_CIERRE_MENSUAL_TIPO_DIARIO '$','$','$','$','$','$','$'"
        Strsql = VBsprintf(Strsql, TxtAno, TxtPeriodo, TxtCod_TipoDiario, vusu, ComputerName(), "S", txtSubDiario)
        
        ExecuteSQL cCONNECT, Strsql
        Unload Me
        
    Case "CANCELAR"
        Unload Me
    End Select
Exit Sub
ERRX:
    Aviso err.Description, 2
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        TxtPeriodo.SetFocus
    End If
End Sub

Private Sub TxtCod_TipoDiario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_TipoDiario("1")
End If
End Sub

Private Sub TxtDes_TipoDiario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_TipoDiario("2")
End If
End Sub

Sub Busca_TipoDiario(Tipo As String)
Dim oTipo As New frmBusqGeneral3
Dim iCol As Long
Dim rstAux As ADODB.Recordset

Set oTipo.oParent = Me

If Tipo = "1" Then
    Strsql = "CN_MUESTRA_SUBDIARIOS_USUARIOS_RESPONSABLES '" & vusu & "','" & Trim(TxtCod_TipoDiario.Text) & "%'"
Else
    Strsql = "CN_MUESTRA_SUBDIARIOS_USUARIOS_RESPONSABLES '" & vusu & "','" & Trim(TxtCod_TipoDiario.Text) & "%'"
End If
With oTipo
    Set .oParent = Me
    .SQuery = Strsql
    .Cargar_Datos
    .Caption = "Selccionar SubDiario"
    codigo = ".."
    Set rstAux = .gexLista.ADORecordset
    
    .gexLista.Columns("Codigo").Width = 700
    .gexLista.Columns("Descripcion").Width = 5000

    
    If rstAux.RecordCount = 1 Then
        codigo = Trim(rstAux!codigo)
        descripcion = Trim(rstAux!descripcion)
    End If
    
    If rstAux.RecordCount > 1 Then .Show vbModal
    
    If codigo <> "" And rstAux.RecordCount > 0 Then
        TxtCod_TipoDiario = codigo
        TxtDes_TipoDiario = descripcion
        
    End If
    TxtAno.SetFocus
End With

codigo = "": descripcion = ""
Unload oTipo
Set oTipo = Nothing
rstAux.Close
Set rstAux = Nothing
End Sub

Sub Busca_TipoSubDiario(Tipo As String)
Dim oTipo As New frmBusqGeneral3
Dim iCol As Long
Dim rstAux As ADODB.Recordset

Set oTipo.oParent = Me

If Tipo = "1" Then
    Strsql = "CN_MUESTRA_SUBDIARIOS_USUARIOS_RESPONSABLES '" & vusu & "','" & Trim(TxtCod_TipoDiario.Text) & "','" & Trim(txtSubDiario.Text) & "','2'"
Else
    Strsql = "CN_MUESTRA_SUBDIARIOS_USUARIOS_RESPONSABLES '" & vusu & "','" & Trim(TxtCod_TipoDiario.Text) & "','" & Trim(txtSubDiario.Text) & "','2'"
End If
With oTipo
    Set .oParent = Me
    .SQuery = Strsql
    .Cargar_Datos
    .Caption = "Selccionar SubDiario"
    codigo = ".."
    Set rstAux = .gexLista.ADORecordset
    
    .gexLista.Columns("Codigo").Width = 700
    .gexLista.Columns("Descripcion").Width = 5000

    
    If rstAux.RecordCount = 1 Then
        codigo = Trim(rstAux!codigo)
        descripcion = Trim(rstAux!descripcion)
    End If
    
    If rstAux.RecordCount > 1 Then .Show vbModal
    
    If codigo <> "" And rstAux.RecordCount > 0 Then
        TxtCod_TipoDiario = codigo
        TxtDes_TipoDiario = descripcion
        
    End If
    TxtAno.SetFocus
End With

codigo = "": descripcion = ""
Unload oTipo
Set oTipo = Nothing
rstAux.Close
Set rstAux = Nothing
End Sub


Private Sub CambiarEstado()
On Error GoTo ERRX

If DevuelveCampo("SELECT COUNT(*) FROM CN_SUBDIARIO WHERE COD_TIPODIARIO ='" & TxtCod_TipoDiario & "'", cCONNECT) > 0 Then
    fraSubdiarios.Visible = True
    Me.txtSubDiario.SetFocus
Else
    Strsql = "CN_CAMBIAR_STATUS_CIERRE_MENSUAL_TIPO_DIARIO '$','$','$','$','$'"
    Strsql = VBsprintf(Strsql, TxtAno, TxtPeriodo, TxtCod_TipoDiario, vusu, ComputerName())
    
    ExecuteSQL cCONNECT, Strsql
    Unload Me
End If

  Exit Sub
ERRX:
    Aviso err.Description, 2
End Sub

Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        TipoAdd = DevuelveCampo("SELECT ISNULL(FLG_STATUS,'N') FROM CN_TipoDiario_Status WHERE ANO = '" & TxtAno & "' AND PERIODO = '" & TxtPeriodo & "' AND COD_TIPODIARIO = '" & TxtCod_TipoDiario & "'", cCONNECT)
        If TipoAdd = "S" Then
            lblStatus = "CERRADO"
            lblStatus.BackColor = &HC000&
        Else
            lblStatus = "PENDIENTE"
            lblStatus.BackColor = &HFF&
        End If
    
        FunctButt1.SetFocus
    End If
End Sub


Private Sub TxtSubDiario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_SubDiario("1")
End If
End Sub

Private Sub TxtDesSubDiario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_SubDiario("2")
End If
End Sub

Sub Busca_SubDiario(Tipo As String)
Dim oTipo As New frmBusqGeneral3
Dim iCol As Long
Dim rstAux As ADODB.Recordset

Set oTipo.oParent = Me

If Tipo = "1" Then
    Strsql = "CN_MUESTRA_SUBDIARIOS_USUARIOS_RESPONSABLES '" & vusu & "','" & Trim(TxtCod_TipoDiario.Text) & "%','" & txtSubDiario & "', 2"
Else
    Strsql = "CN_MUESTRA_SUBDIARIOS_USUARIOS_RESPONSABLES '" & vusu & "','" & Trim(TxtCod_TipoDiario.Text) & "%','" & txtSubDiario & "',2"
End If
With oTipo
    Set .oParent = Me
    .SQuery = Strsql
    .Cargar_Datos
    .Caption = "Selccionar SubDiario"
    codigo = ".."
    Set rstAux = .gexLista.ADORecordset
    
    .gexLista.Columns("Codigo").Width = 700
    .gexLista.Columns("Descripcion").Width = 5000

    
    If rstAux.RecordCount = 1 Then
        codigo = Trim(rstAux!codigo)
        descripcion = Trim(rstAux!descripcion)
    End If
    
    If rstAux.RecordCount > 1 Then .Show vbModal
    
    If codigo <> "" And rstAux.RecordCount > 0 Then
        txtSubDiario = codigo
        txtDesSubdiario = descripcion
        
        TipoAdd = DevuelveCampo("SELECT ISNULL(FLG_STATUS,'N') FROM CN_TipoDiario_Status WHERE ANO = '" & TxtAno & "' AND PERIODO = '" & TxtPeriodo & "' AND COD_TIPODIARIO = '" & TxtCod_TipoDiario & "' AND Cod_SubDiario ='" & txtSubDiario & "'", cCONNECT)
        If TipoAdd = "S" Then
            lblSDStatus = "CERRADO"
            lblSDStatus.BackColor = &HC000&
        Else
            lblSDStatus = "PENDIENTE"
            lblSDStatus.BackColor = &HFF&
        End If
        
        
    End If
    FunctButt2.SetFocus
End With

codigo = "": descripcion = ""
Unload oTipo
Set oTipo = Nothing
rstAux.Close
Set rstAux = Nothing
End Sub

