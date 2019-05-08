VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmImportaciones 
   Caption         =   "Generación de Num. Importación"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDatos 
      Height          =   5760
      Left            =   45
      TabIndex        =   1
      Top             =   -30
      Width           =   7215
      Begin VB.TextBox TxtDes_OrigenImportacion 
         Height          =   330
         Left            =   3240
         TabIndex        =   19
         Top             =   4140
         Width           =   3375
      End
      Begin VB.TextBox TxtDes_Embarque 
         Height          =   330
         Left            =   3240
         TabIndex        =   18
         Top             =   3660
         Width           =   3375
      End
      Begin VB.TextBox TxtDes_Proveedor 
         Height          =   330
         Left            =   3240
         TabIndex        =   17
         Top             =   3180
         Width           =   3855
      End
      Begin VB.TextBox TxtObservacion 
         Height          =   930
         Left            =   2040
         ScrollBars      =   1  'Horizontal
         TabIndex        =   16
         Top             =   4620
         Width           =   5055
      End
      Begin VB.TextBox TxtCod_OrigenImportacion 
         Height          =   330
         Left            =   2040
         TabIndex        =   15
         Top             =   4140
         Width           =   1215
      End
      Begin VB.TextBox TxtCod_Embarque 
         Height          =   330
         Left            =   2040
         TabIndex        =   14
         Top             =   3660
         Width           =   1215
      End
      Begin VB.TextBox Txtcod_Proveedor 
         Height          =   330
         Left            =   2040
         TabIndex        =   13
         Top             =   3180
         Width           =   1215
      End
      Begin VB.TextBox TxtNum_guiaAerea 
         Height          =   330
         Left            =   2040
         TabIndex        =   12
         Top             =   2220
         Width           =   1695
      End
      Begin VB.TextBox Txtcod_Ordcomp 
         Height          =   330
         Left            =   3000
         TabIndex        =   11
         Top             =   1740
         Width           =   1335
      End
      Begin VB.TextBox TxtSerOrdComp 
         Height          =   330
         Left            =   2040
         TabIndex        =   10
         Top             =   1740
         Width           =   855
      End
      Begin VB.TextBox TxtNumImportacion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdOrigenImportacion 
         Caption         =   "..."
         Height          =   330
         Left            =   6720
         TabIndex        =   8
         Top             =   4140
         Width           =   375
      End
      Begin VB.CommandButton CmdEmbarque 
         Caption         =   "..."
         Height          =   330
         Left            =   6720
         TabIndex        =   7
         Top             =   3660
         Width           =   375
      End
      Begin VB.TextBox TxtTipoAnexo 
         Height          =   350
         Left            =   2055
         TabIndex        =   6
         Text            =   "P"
         Top             =   750
         Width           =   495
      End
      Begin VB.TextBox Txtdes_Anexo 
         Height          =   350
         Left            =   3195
         TabIndex        =   5
         Top             =   765
         Width           =   3915
      End
      Begin VB.TextBox TxtCod_Anexo 
         Height          =   350
         Left            =   2580
         TabIndex        =   4
         Top             =   750
         Width           =   570
      End
      Begin VB.TextBox txtCod_TipPro 
         Height          =   330
         Left            =   2055
         TabIndex        =   3
         Top             =   1245
         Width           =   525
      End
      Begin VB.TextBox txtDes_TipPro 
         Height          =   330
         Left            =   3015
         TabIndex        =   2
         Top             =   1245
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTFechaGuia 
         Height          =   345
         Left            =   2040
         TabIndex        =   20
         Top             =   2700
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         Format          =   23920641
         CurrentDate     =   38110
      End
      Begin MSComCtl2.DTPicker dtpFec_RegDoc 
         Height          =   345
         Left            =   5595
         TabIndex        =   21
         Top             =   1740
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   609
         _Version        =   393216
         Format          =   23920641
         CurrentDate     =   38110
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Observación"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   4740
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Origen Importación"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   4260
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Embarque"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   3825
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prov. Agente Importación"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   3300
         Width           =   1800
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Guia Aerea"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   2820
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Num. Guia Aerea"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   2340
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orden Compra"
         Height          =   195
         Left            =   255
         TabIndex        =   26
         Top             =   1815
         Width           =   1020
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Num_Importación"
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
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Anexo Contable:"
         Height          =   195
         Left            =   255
         TabIndex        =   24
         Top             =   825
         Width           =   1170
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Procedencia"
         Height          =   195
         Left            =   270
         TabIndex        =   23
         Top             =   1305
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Registro"
         Height          =   345
         Left            =   4425
         TabIndex        =   22
         Top             =   1815
         Width           =   1095
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2400
      TabIndex        =   0
      Top             =   5865
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "0~0~ACEPTAR~True~True~&Aceptar~0~0~1~~0~False~False~&Aceptar~~1~0~CANCELAR~True~True~&Cancelar~0~0~2~~0~False~False~&Cancelar~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmImportaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Accion As String
Dim StrSQL As String
Public Codigo, Descripcion  As String
Public oParent As Object

Private Sub CmdEmbarque_Click()
Load FrmManteEmbarques
With FrmManteEmbarques
    .Show 1
End With
Set FrmManteEmbarques = Nothing
End Sub

Private Sub CmdOrigenImportacion_Click()
Load FrmManteOrigenImportacion
With FrmManteOrigenImportacion
    .Show 1
End With
Set FrmManteOrigenImportacion = Nothing
End Sub

Private Sub DTFechaGuia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    If SALVAR_DATOS = True Then Unload Me
Case "CANCELAR"
    Unload Me
End Select
End Sub

Function SALVAR_DATOS() As Boolean
Dim Pregunta As Variant
Dim mRs As ADODB.Recordset

On Error GoTo errSalvarDatos

SALVAR_DATOS = True

If UCase(Me.Accion) = "C" Then Exit Function
If UCase(Me.Accion) = "D" Then
    Pregunta = MsgBox("¿Está seguro de eliminar el registro?", vbYesNo)
    If Pregunta = vbNo Then Exit Function
End If

StrSQL = "UP_MAN_LG_IMPORTACIONES '" & "I" & "','" & TxtNumImportacion.Text & "','" & Trim(TxtSerOrdComp.Text) & "','" & Trim(Me.Txtcod_Ordcomp.Text) & "','" & _
            TxtNum_guiaAerea & "','" & DTFechaGuia.Value & "','" & Txtcod_Proveedor.Text & "','" & _
            TxtCod_Embarque & "','" & TxtCod_OrigenImportacion.Text & "','" & TxtObservacion.Text & "','" & TxtTipoAnexo.Text & "','" & TxtCod_Anexo.Text & "','" & txtCod_TipPro.Text & "','" & dtpFec_RegDoc.Value & "'"
        
Set mRs = GetRecordset(cConnect, StrSQL)
If Not mRs Is Nothing Then
    Aviso "Número de Importación Generado: " & mRs!Num_Importacion, 2
    mRs.Close
    Set mRs = Nothing
End If

'Mensaje kMESSAGE_INF_PROCESS_SATISFACTO

Exit Function
errSalvarDatos:
    SALVAR_DATOS = False
    ErrorHandler Err, "SALVAR_DATOS"
End Function

Sub MUESTRA_EMBARQUES(Tipo As Integer)
    Select Case Tipo
        Case 1:
                StrSQL = "SELECT des_Embarque from TG_TIPEMB where cod_embarque = '" & TxtCod_Embarque.Text & "'"
                TxtDes_Embarque.Text = Trim(DevuelveCampo(StrSQL, cConnect))
                TxtCod_OrigenImportacion.SetFocus
        Case 2, 3:
                Dim oTipo As New frmBusqGeneral
                Dim Rs As New ADODB.Recordset
                Set oTipo.oParent = Me
                
                If Tipo = 2 Then
                    oTipo.sQuery = "SELECT cod_Embarque as Código, Des_embarque as Descripción FROM TG_TIPEMB WHERE Des_Embarque LIKE '" & Trim(TxtDes_Embarque.Text) & "%'"
                Else
                    oTipo.sQuery = "SELECT cod_Embarque as Código, Des_embarque as Descripción FROM TG_TIPEMB"
                End If
                
                oTipo.CARGAR_DATOS
                oTipo.Show 1
                If Codigo <> "" Then
                    TxtCod_Embarque.Text = Trim(Codigo)
                    TxtDes_Embarque.Text = Trim(Descripcion)
                    TxtCod_OrigenImportacion.SetFocus
                End If
                Codigo = "": Descripcion = ""
                Set oTipo = Nothing
                Set Rs = Nothing
    End Select
End Sub

Sub MUESTRA_ORIGENIMPORTACION(Tipo As Integer)
    Select Case Tipo
        Case 1:
                StrSQL = "SELECT des_OrigenImportacion  from LG_ORIGEN_IMPORTACION where cod_OrigenImportacion = '" & TxtCod_OrigenImportacion.Text & "'"
                TxtDes_OrigenImportacion.Text = Trim(DevuelveCampo(StrSQL, cConnect))
        Case 2, 3:
                Dim oTipo As New frmBusqGeneral
                Dim Rs As New ADODB.Recordset
                Set oTipo.oParent = Me
                
                If Tipo = 2 Then
                    oTipo.sQuery = "SELECT cod_OrigenImportacion as Código, Des_OrigenImportacion as Descripción FROM LG_ORIGEN_IMPORTACION WHERE Des_OrigenImportacion LIKE '" & Trim(TxtDes_OrigenImportacion.Text) & "%'"
                Else
                    oTipo.sQuery = "SELECT cod_OrigenImportacion as Código, Des_OrigenImportacion as Descripción FROM LG_ORIGEN_IMPORTACION"
                End If
                
                oTipo.CARGAR_DATOS
                oTipo.Show 1
                If Codigo <> "" Then
                    TxtCod_OrigenImportacion.Text = Trim(Codigo)
                    TxtDes_OrigenImportacion.Text = Trim(Descripcion)
                    TxtObservacion.SetFocus
                End If
                Codigo = "": Descripcion = ""
                Set oTipo = Nothing
                Set Rs = Nothing
    End Select
End Sub

Sub MUESTRA_PROVEEDOR(Tipo As Integer)
    Select Case Tipo
        Case 1:
                StrSQL = "SELECT des_Proveedor from LG_PROVEEDOR where cod_Proveedor = '" & Txtcod_Proveedor.Text & "'"
                TxtDes_Proveedor.Text = Trim(DevuelveCampo(StrSQL, cConnect))
                TxtCod_Embarque.SetFocus
        Case 2, 3:
                Dim oTipo As New frmBusqGeneral
                Dim Rs As New ADODB.Recordset
                Set oTipo.oParent = Me
                
                If Tipo = 2 Then
                    oTipo.sQuery = "SELECT cod_Proveedor as Código, Des_Proveedor as Descripción FROM LG_PROVEEDOR WHERE Des_Proveedor LIKE '%" & Trim(TxtDes_Proveedor.Text) & "%'"
                Else
                    oTipo.sQuery = "SELECT cod_Proveedor as Código, Des_Proveedor as Descripción FROM LG_PROVEEDOR"
                End If
                
                oTipo.CARGAR_DATOS
                oTipo.Show 1
                If Codigo <> "" Then
                    Txtcod_Proveedor.Text = Trim(Codigo)
                    TxtDes_Proveedor.Text = Trim(Descripcion)
                    TxtCod_Embarque.SetFocus
                End If
                Codigo = "": Descripcion = ""
                
                Set oTipo = Nothing
                Set Rs = Nothing
    End Select
End Sub

Private Sub Txtcod_Ordcomp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TxtSerOrdComp.Text <> "" Then
        Call BUSCA_NUM_ORDCOMP
    End If
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtObservacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtCod_Embarque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxtCod_Embarque.Text) = "" Then
         Call MUESTRA_EMBARQUES(3)
    Else
        Call MUESTRA_EMBARQUES(1)
    End If
End If
End Sub


Private Sub TxtCod_OrigenImportacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxtCod_Embarque.Text) = "" Then
         Call MUESTRA_ORIGENIMPORTACION(3)
    Else
        If TxtCod_OrigenImportacion.Text = "" Then
            Call MUESTRA_ORIGENIMPORTACION(2)
        Else
            Call MUESTRA_ORIGENIMPORTACION(1)
        End If
    End If
End If
End Sub

Private Sub Txtcod_Proveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Txtcod_Proveedor.Text) = "" Then
         Call MUESTRA_PROVEEDOR(3)
    Else
        Call MUESTRA_PROVEEDOR(1)
    End If
End If
End Sub

Private Sub TxtDes_Embarque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call MUESTRA_EMBARQUES(2)
End If
End Sub

Private Sub TxtDes_OrigenImportacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call MUESTRA_ORIGENIMPORTACION(2)
End If
End Sub

Private Sub TxtDes_Proveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call MUESTRA_PROVEEDOR(2)
End If
End Sub

Private Sub TxtNum_guiaAerea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtSerOrdComp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxtSerOrdComp.Text) = "" Then
        SendKeys "{TAB}"
    Else
        TxtSerOrdComp.Text = Right("000" & Trim(TxtSerOrdComp.Text), 3)
    End If
    
End If
End Sub

Sub BUSCA_SERIE()
Dim oTipo As New frmBusqGeneral
Dim Rs As New ADODB.Recordset
Set oTipo.oParent = Me

oTipo.sQuery = "SELECT DISTINCT Ser_Ordcomp as Serie FROM LG_ORDCOMP"

oTipo.CARGAR_DATOS
oTipo.Show 1
If Codigo <> "" Then
    TxtSerOrdComp.Text = Trim(Codigo)
    Txtcod_Ordcomp.SetFocus
End If
Codigo = "": Descripcion = ""

Set oTipo = Nothing
Set Rs = Nothing
End Sub

Sub BUSCA_NUM_ORDCOMP()
Dim oTipo As New frmBusqGeneral
Dim Rs As New ADODB.Recordset
Set oTipo.oParent = Me

Txtcod_Ordcomp.Text = StrZero(Txtcod_Ordcomp.Text, 6)
oTipo.sQuery = "SELECT Cod_OrdComp , Cod_Proveedor FROM LG_ORDCOMP where Ser_OrdComp ='" & Trim(TxtSerOrdComp.Text) & "' AND Cod_OrdComp = '" & Txtcod_Ordcomp.Text & "'"

oTipo.CARGAR_DATOS
oTipo.DGridLista.Columns(1).Width = 0
oTipo.Show 1
If Codigo <> "" Then
    Txtcod_Ordcomp.Text = Trim(Codigo)
    Txtcod_Proveedor.Text = Trim(Descripcion)
    Txtcod_Ordcomp.SetFocus
End If
Codigo = "": Descripcion = ""

Set oTipo = Nothing
Set Rs = Nothing
End Sub


Private Sub TxtCod_Anexo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxtCod_Anexo.Text) = "" Then
        Txtdes_Anexo.SetFocus
    Else
        Txtdes_Anexo.Text = DevuelveCampo("select des_anexo from cn_anexoscontables where cod_tipAnex='" & Trim(TxtTipoAnexo.Text) & "' and cod_anxo='" & Trim(TxtCod_Anexo.Text) & "'", cConnect)
        SendKeys "{TAB}"
    End If
End If
End Sub

Private Sub Txtdes_Anexo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call MUESTRA_ANEXOS(2)
End If
End Sub

Private Sub TxtTipoAnexo_Change()
TxtTipoAnexo.Text = UCase(TxtTipoAnexo.Text)
End Sub

Private Sub TxtTipoAnexo_KeyPress(KeyAscii As Integer)
Dim Conta As Integer
If KeyAscii = 13 Then
    If Trim(TxtTipoAnexo.Text) = "" Then
        Call MUESTRA_TIPOANEXOS
    Else
        Conta = DevuelveCampo("select count(*) from CN_TipoAnexoContable where cod_tipanex='" & Trim(TxtTipoAnexo.Text) & "'", cConnect)
        If Conta = 0 Then
            MsgBox "Tipo Anexo no valido", vbCritical
            SelectionText TxtTipoAnexo
            Exit Sub
        End If
        SendKeys "{TAB}"
    End If
End If
End Sub

Sub MUESTRA_TIPOANEXOS()
Dim oTipo As New frmBusqGeneral
Dim Rs As New ADODB.Recordset
Set oTipo.oParent = Me

oTipo.sQuery = "SELECT cod_TipAnex as Código, Des_TipAnex as Descripción FROM CN_TipoAnexoContable"

oTipo.CARGAR_DATOS
oTipo.Show 1
If Codigo <> "" Then
    TxtTipoAnexo.Text = Trim(Codigo)
    SendKeys "{TAB}"
End If
Codigo = "": Descripcion = ""

Set oTipo = Nothing
Set Rs = Nothing
End Sub


Sub MUESTRA_ANEXOS(ByVal Tipo As Integer)
Dim oTipo As New frmBusqGeneral
Dim Rs As New ADODB.Recordset
Set oTipo.oParent = Me

If Tipo = 1 Then
    oTipo.sQuery = "SELECT cod_Anxo as Código, Des_Anexo as Descripción FROM CN_AnexoSContables WHERE COD_TIPANEX='" & Trim(TxtTipoAnexo.Text) & "'"
Else
    oTipo.sQuery = "SELECT cod_Anxo as Código, Des_Anexo as Descripción FROM CN_AnexoSContables WHERE COD_TIPANEX='" & Trim(TxtTipoAnexo.Text) & "' AND DES_ANEXO LIKE '%" & Trim(Txtdes_Anexo.Text) & "%' order by des_anexo"
End If
oTipo.CARGAR_DATOS
oTipo.Show 1
If Codigo <> "" Then
    TxtCod_Anexo.Text = Trim(Codigo)
    Txtdes_Anexo.Text = Trim(Descripcion)
    If txtCod_TipPro.Enabled Then
        txtCod_TipPro.SetFocus
    End If
End If
Codigo = "": Descripcion = ""

Set oTipo = Nothing
Set Rs = Nothing
End Sub



Sub MUESTRA_PROCEDENCIA(Tipo As Integer)
    Select Case Tipo
        Case 1:
                StrSQL = "SELECT des_TipPro from CN_PROCEDENCIA where cod_TIPPRO = '" & txtCod_TipPro.Text & "'"
                txtDes_TipPro.Text = Trim(DevuelveCampo(StrSQL, cConnect))
                If TxtSerOrdComp.Enabled Then
                    TxtSerOrdComp.SetFocus
                Else
                    TxtNum_guiaAerea.SetFocus
                End If
        Case 2, 3:
                Dim oTipo As New frmBusqGeneral
                Dim Rs As New ADODB.Recordset
                Set oTipo.oParent = Me
                
                If Tipo = 2 Then
                    oTipo.sQuery = "SELECT cod_tIPPRO as Código, Des_TipPro as Descripción FROM CN_PROCEDENCIA WHERE Des_TipPro LIKE '" & Trim(txtDes_TipPro.Text) & "%'"
                Else
                    oTipo.sQuery = "SELECT cod_tIPPRO as Código, Des_TipPro as Descripción FROM CN_PROCEDENCIA "
                End If
                
                oTipo.CARGAR_DATOS
                oTipo.Show 1
                If Codigo <> "" Then
                    txtCod_TipPro.Text = Trim(Codigo)
                    txtDes_TipPro.Text = Trim(Descripcion)
                    If TxtSerOrdComp.Enabled Then
                        TxtSerOrdComp.SetFocus
                    Else
                        TxtNum_guiaAerea.SetFocus
                    End If
                End If
                Codigo = "": Descripcion = ""
                Set oTipo = Nothing
                Set Rs = Nothing
    End Select
End Sub


Private Sub TxtCod_TipPro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(txtCod_TipPro.Text) = "" Then
         Call MUESTRA_PROCEDENCIA(3)
    Else
        Call MUESTRA_PROCEDENCIA(1)
    End If
End If
End Sub



Public Sub Aviso(Mensaje As String, Tipo As Integer)
    Select Case Tipo
        Case 1
            MsgBox Mensaje, vbExclamation, "Aviso"
        Case 2
            MsgBox Mensaje, vbInformation + vbMsgBoxRight, "Mensaje"
        Case 3
            MsgBox Mensaje, vbCritical, "Error Grave"
    End Select
End Sub

