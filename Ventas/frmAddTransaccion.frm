VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAddTransaccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transmitir"
   ClientHeight    =   2820
   ClientLeft      =   2430
   ClientTop       =   1470
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frTransacciones 
      BorderStyle     =   0  'None
      Height          =   3090
      Left            =   0
      TabIndex        =   5
      Top             =   -240
      Width           =   6885
      Begin VB.Frame Frame3 
         Height          =   2250
         Left            =   0
         TabIndex        =   3
         Top             =   180
         Width           =   6795
         Begin VB.TextBox TxtDes_Retencion 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   2220
            TabIndex        =   16
            Top             =   1860
            Width           =   3930
         End
         Begin VB.TextBox TxtTipo_Retencion 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   1620
            MaxLength       =   10
            TabIndex        =   15
            Top             =   1860
            Width           =   585
         End
         Begin VB.TextBox TxtPorcentaje_Retencion 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Left            =   6165
            MaxLength       =   10
            TabIndex        =   14
            Top             =   1860
            Width           =   480
         End
         Begin VB.CheckBox chkDetraccion 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Caption         =   "Pagar Detraccion :"
            Height          =   255
            Left            =   135
            TabIndex        =   12
            Top             =   1590
            Width           =   1680
         End
         Begin VB.TextBox TxtCod_AreaResponsable 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   1620
            MaxLength       =   10
            TabIndex        =   11
            Top             =   1260
            Width           =   600
         End
         Begin VB.TextBox TxtDes_AreaResponsable 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   2220
            TabIndex        =   10
            Top             =   1260
            Width           =   4425
         End
         Begin VB.TextBox txtDes_TipIte 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   2220
            TabIndex        =   8
            Top             =   510
            Width           =   4425
         End
         Begin VB.TextBox txtCod_TipIte 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   1620
            MaxLength       =   4
            TabIndex        =   0
            Top             =   510
            Width           =   600
         End
         Begin VB.TextBox TxtObservacion 
            Height          =   465
            Left            =   1620
            TabIndex        =   1
            Top             =   795
            Width           =   5025
         End
         Begin MSComCtl2.DTPicker DTPFecha 
            Height          =   300
            Left            =   1620
            TabIndex        =   4
            Top             =   210
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   529
            _Version        =   393216
            Format          =   94109697
            CurrentDate     =   38590
         End
         Begin VB.Label lbTipoDetraccion 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Detraccion:"
            Height          =   195
            Left            =   135
            TabIndex        =   17
            Top             =   1935
            Width           =   1185
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Area Responsable:"
            Height          =   195
            Left            =   135
            TabIndex        =   13
            Top             =   1305
            Width           =   1350
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Item :"
            Height          =   195
            Left            =   135
            TabIndex        =   9
            Top             =   555
            Width           =   975
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Glosa :"
            Height          =   195
            Left            =   135
            TabIndex        =   7
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Registro :"
            Height          =   195
            Left            =   135
            TabIndex        =   6
            Top             =   263
            Width           =   1395
         End
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   4290
         TabIndex        =   2
         Top             =   2490
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmAddTransaccion.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
End
Attribute VB_Name = "frmAddTransaccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codigo As String, Descripcion As String, TipoAdd As String, scorrelativo As String, lfSalvar As Boolean

Public sTipoBusq As String, Tipoa As String, Tipob As String, sNum_Planilla_Letra As Integer
Dim RsGrid1 As ADODB.Recordset
Dim RsGrid2 As ADODB.Recordset
Dim intTransaccion As Integer, vrTotalTransaccion As Double
Dim strSQL As String, intCancel As Integer

Private Sub Form_Load()
'txtCod_TipIte.SetFocus
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim sMessage As Long
Dim strSQL As String

On Error GoTo dprError

Select Case ActionName
  Case "GRABAR"
       Dim SQL As String
       Dim sDetraccion  As String, Cod_AreaResponsable  As String, Tipo_Retencion  As String
            sDetraccion = IIf(chkDetraccion, "S", "N")
            Cod_AreaResponsable = FixNulos(TxtCod_AreaResponsable.Text, vbString)
            Tipo_Retencion = FixNulos(TxtTipo_Retencion.Text, vbString)
            
             SQL = " exec CN_ENVIA_FACTURA_VENTAS_A_INKADESIGNS '" & scorrelativo & "' , '" & DTPFecha.Value & "', '" & txtCod_TipIte.Text & "','" & TxtObservacion.Text & "','" & Cod_AreaResponsable & "','" & vusu & "','" & sDetraccion & "','" & Tipo_Retencion & "'"
                     ExecuteSQL cCONNECT, SQL
                     Unload Me
   

  Case "CANCELAR"
      Unload Me
    
End Select

Exit Sub

dprError:

errores err.Number

End Sub


Private Sub txtCod_TipIte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_TipoItem(1)
End If

End Sub


Sub Busca_TipoItem(Tipo As Integer)
Dim oTipo As New frmBusqGeneral3
Dim RS As Object
Set RS = CreateObject("ADODB.Recordset")

Set oTipo.oParent = Me
If Tipo = 1 Then
    oTipo.SQuery = "select Cod_TipIte as Codigo, Des_TipIte as Descripcion,por_retencion as PorcRetencion from INKADESIGNS..cn_tipoitem where Cod_TipIte  like '%" & Trim(txtCod_TipIte.Text) & "%'"
Else
    oTipo.SQuery = "select Cod_TipIte as Codigo, Des_TipIte as Descripcion,,por_retencion as PorcRetencion from INKADESIGNS..cn_tipoitem where Des_TipIte  like '%" & Trim(txtDes_TipIte.Text) & "%'"
End If

oTipo.Caption = "Buscar Tipo Item"
oTipo.CARGAR_DATOS

oTipo.gexLista.Columns("Codigo").Width = 900
oTipo.gexLista.Columns("Descripcion").Width = 2500
oTipo.gexLista.Columns("PorcRetencion").Width = 800


If oTipo.gexLista.RowCount > 1 Then
    oTipo.Show vbModal
Else
    codigo = oTipo.gexLista.Value(oTipo.gexLista.Columns("Codigo").Index)
    Descripcion = oTipo.gexLista.Value(oTipo.gexLista.Columns("Descripcion").Index)
    TipoAdd = oTipo.gexLista.Value(oTipo.gexLista.Columns("PorcRetencion").Index)
End If

If Trim(codigo) <> "" Then
    txtCod_TipIte.Text = codigo
    txtDes_TipIte.Text = Descripcion
    'por_RetencionRentaVarios = TipoAdd
    codigo = "": Descripcion = "": TipoAdd = ""
    TxtObservacion.SetFocus
    
End If
Unload oTipo
Set oTipo = Nothing
Set RS = Nothing
End Sub


Private Sub txtDes_TipIte_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        TxtObservacion.SetFocus
    End If

End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then TxtCod_AreaResponsable.SetFocus
End Sub

'/*/*/*/*/*

Private Sub TxtCod_AreaResponsable_KeyPress(KeyAscii As Integer)
TxtDes_AreaResponsable.Text = ""

If KeyAscii = 13 Then
    If Trim(TxtCod_AreaResponsable.Text) = "" Then
        Call Busca_AreaResponsable(1)
    Else
        TxtDes_AreaResponsable.Text = DevuelveCampo("select des_arearesponsable from INKADESIGNS..cn_areasresponsables where cod_arearesponsable ='" & TxtCod_AreaResponsable.Text & "' and flg_estadoUbicacion='A'", cCONNECT)
        If RTrim(TxtDes_AreaResponsable.Text) = "" Then
            Aviso "AREA RESPONSABLE INEXISTENTE", 2
            TxtCod_AreaResponsable.SetFocus
            Exit Sub
        End If
    End If
End If
End Sub

Private Sub TxtDes_AreaResponsable_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_AreaResponsable(2)
End If
End Sub

Sub Busca_AreaResponsable(Tipo As Integer)
Dim oTipo As New frmBusqGeneral3
Dim RS As Object
Set RS = CreateObject("ADODB.Recordset")

Set oTipo.oParent = Me
If Tipo = 1 Then
    oTipo.SQuery = "select Cod_AreaResponsable as Codigo, Des_AreaResponsable as Descripcion  from INKADESIGNS..CN_AreasResponsables where flg_estadoUbicacion='A'"
Else
    oTipo.SQuery = "select Cod_AreaResponsable as Codigo, Des_AreaResponsable as Descripcion  from INKADESIGNS..CN_AreasResponsables where des_arearesponsable like '%" & Trim(TxtDes_AreaResponsable.Text) & "%' and flg_estadoUbicacion='A'"
End If

oTipo.Caption = "Buscar Area"
oTipo.CARGAR_DATOS

oTipo.gexLista.Columns("Codigo").Width = 900
oTipo.gexLista.Columns("Descripcion").Width = 2500

If oTipo.gexLista.RowCount > 1 Then
    oTipo.Show vbModal
Else
    codigo = oTipo.gexLista.Value(oTipo.gexLista.Columns("Codigo").Index)
    Descripcion = oTipo.gexLista.Value(oTipo.gexLista.Columns("Descripcion").Index)
End If

If Trim(codigo) <> "" Then
    TxtCod_AreaResponsable = codigo 'oTipo.gexLista.Value(oTipo.gexLista.Columns("Codigo").Index)
    TxtDes_AreaResponsable.Text = Descripcion 'oTipo.gexLista.Value(oTipo.gexLista.Columns("Descripcion").Index)
    codigo = "": Descripcion = "": TipoAdd = ""
    SendKeys "{TAB}"
End If

Unload oTipo
Set oTipo = Nothing
Set RS = Nothing
End Sub


Private Sub TxtTipo_Retencion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_TipoRetencion("1")
End If
End Sub
Private Sub TxtDes_Retencion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_TipoRetencion("2")
End If
End Sub

Sub Busca_TipoRetencion(Tipo As Integer)
Dim oTipo As New frmBusqGeneral3
Dim RS As Object
Set RS = CreateObject("ADODB.Recordset")

Set oTipo.oParent = Me
If Tipo = 1 Then
    oTipo.SQuery = "select tipo_Retencion as Codigo, Des_Retencion as Descripcion, Porcentaje ,cod_retencion,  Anexo, Numeral from cn_TipoRetencion where tipo_retencion like '%" & Trim(TxtTipo_Retencion.Text) & "%' AND FLG_SELECCIONABLE ='S'"
Else
    oTipo.SQuery = "select tipo_Retencion as Codigo, Des_Retencion as Descripcion, Porcentaje ,cod_retencion,  Anexo, Numeral  from cn_TipoRetencion where Des_Retencion like '%" & Trim(TxtDes_Retencion.Text) & "%' AND FLG_SELECCIONABLE ='S'"
End If

oTipo.Caption = "Buscar Tipo Detraccion"
oTipo.CARGAR_DATOS

oTipo.gexLista.Columns("Codigo").Width = 900
oTipo.gexLista.Columns("cod_retencion").Width = 900
oTipo.gexLista.Columns("Descripcion").Width = 2500
oTipo.gexLista.Columns("Porcentaje").Width = 900

If oTipo.gexLista.RowCount > 1 Then
    oTipo.Show vbModal
Else
    codigo = oTipo.gexLista.Value(oTipo.gexLista.Columns("Codigo").Index)
    Descripcion = oTipo.gexLista.Value(oTipo.gexLista.Columns("Descripcion").Index)
    TipoAdd = oTipo.gexLista.Value(oTipo.gexLista.Columns("Porcentaje").Index)
End If

If Trim(codigo) <> "" Then
    TxtTipo_Retencion.Text = codigo 'oTipo.gexLista.Value(oTipo.gexLista.Columns("Codigo").Index)
    TxtDes_Retencion.Text = Descripcion 'oTipo.gexLista.Value(oTipo.gexLista.Columns("Descripcion").Index)
    TxtPorcentaje_Retencion.Text = TipoAdd 'oTipo.gexLista.Value(oTipo.gexLista.Columns("Porcentaje").Index)
    codigo = "": Descripcion = "": TipoAdd = ""
    SendKeys "{TAB}"
End If
Unload oTipo
Set oTipo = Nothing
Set RS = Nothing
End Sub

