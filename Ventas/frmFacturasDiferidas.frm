VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmFacturasDiferidas 
   Caption         =   "Facturas Diferidas"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatos 
      Height          =   1080
      Left            =   45
      TabIndex        =   13
      Top             =   585
      Width           =   9045
      Begin VB.TextBox txtRuc 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   960
         MaxLength       =   11
         TabIndex        =   0
         Top             =   315
         Width           =   1200
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   2355
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "C"
         Top             =   330
         Width           =   360
      End
      Begin VB.TextBox txtDes_Anexo 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   3330
         MaxLength       =   30
         TabIndex        =   3
         Top             =   330
         Width           =   4155
      End
      Begin VB.TextBox txtCod_Anexo 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   2730
         MaxLength       =   4
         TabIndex        =   2
         Top             =   330
         Width           =   600
      End
      Begin NumBoxProject.NumBox inpFec_Inicio 
         Height          =   285
         Left            =   1710
         TabIndex        =   4
         Top             =   675
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   2
         Text            =   ""
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   0
      End
      Begin VB.OptionButton optCliente 
         Caption         =   "Cliente"
         Height          =   330
         Left            =   105
         TabIndex        =   15
         Top             =   315
         Value           =   -1  'True
         Width           =   1605
      End
      Begin VB.OptionButton optFecha 
         Caption         =   "Fecha de Emisión"
         Height          =   330
         Left            =   105
         TabIndex        =   14
         Top             =   660
         Width           =   1755
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   435
         Left            =   7500
         TabIndex        =   6
         Top             =   225
         Width           =   1395
      End
      Begin NumBoxProject.NumBox inpFec_Final 
         Height          =   285
         Left            =   3015
         TabIndex        =   5
         Top             =   675
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   2
         Text            =   ""
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   0
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2220
         TabIndex        =   18
         Tag             =   "Anexo Type"
         Top             =   345
         Width           =   90
      End
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   45
      TabIndex        =   7
      Top             =   0
      Width           =   9045
      Begin VB.TextBox txtDes_Origen 
         Height          =   285
         Left            =   1245
         TabIndex        =   11
         Top             =   210
         Width           =   1575
      End
      Begin VB.TextBox txtCod_Origen 
         Height          =   285
         Left            =   810
         MaxLength       =   1
         TabIndex        =   10
         Text            =   "N"
         Top             =   210
         Width           =   375
      End
      Begin VB.OptionButton optDiferidos 
         Caption         =   "Pendientes"
         Height          =   210
         Left            =   2985
         TabIndex        =   9
         Top             =   225
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.OptionButton opt 
         Caption         =   "Canceladas"
         Height          =   210
         Left            =   4635
         TabIndex        =   8
         Top             =   225
         Width           =   1410
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Origen :"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   255
         Width           =   495
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5340
      Left            =   30
      TabIndex        =   16
      Top             =   1755
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   9419
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmFacturasDiferidas.frx":0000
      Column(2)       =   "frmFacturasDiferidas.frx":00C8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmFacturasDiferidas.frx":016C
      FormatStyle(2)  =   "frmFacturasDiferidas.frx":02A4
      FormatStyle(3)  =   "frmFacturasDiferidas.frx":0354
      FormatStyle(4)  =   "frmFacturasDiferidas.frx":0408
      FormatStyle(5)  =   "frmFacturasDiferidas.frx":04E0
      FormatStyle(6)  =   "frmFacturasDiferidas.frx":0598
      FormatStyle(7)  =   "frmFacturasDiferidas.frx":0678
      FormatStyle(8)  =   "frmFacturasDiferidas.frx":0724
      ImageCount      =   0
      PrinterProperties=   "frmFacturasDiferidas.frx":07D4
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3345
      TabIndex        =   17
      Top             =   7215
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmFacturasDiferidas.frx":09AC
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   7140
      Top             =   7365
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmFacturasDiferidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sOpcion As String
Public codigo As String, Descripcion As String
Dim strSQL  As String

Private Sub cmdBuscar_Click()
    Buscar
End Sub

Private Sub Form_Load()
    sOpcion = "1"
End Sub


Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "CAMBIOESTADO"
            If gridex1.RowCount = 0 Then Exit Sub
            CambioEstado
        Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub inpFec_Final_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub inpFec_Inicio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub


Private Sub optFecha_Click()
    sOpcion = "3"
    inpFec_Inicio.SetFocus
End Sub



Sub Buscar()
Dim strSQL As String
Dim sFlg_Pendiente_Cancelado As String
On Error GoTo errores

If txtCod_Anexo.Text <> "" Then
    sOpcion = "2"
Else
    If optFecha Then
        sOpcion = "3"
    Else
        sOpcion = "1"
    End If
End If

If optDiferidos Then
    sFlg_Pendiente_Cancelado = "P"
Else
    sFlg_Pendiente_Cancelado = "C"
End If

strSQL = "CN_VENTAS_MUESTRA_FACTURAS_DIFERIDAS '$','$','$','$' , '$', '$' ,'$'"
strSQL = VBsprintf(strSQL, sOpcion, sFlg_Pendiente_Cancelado, txtCod_Origen.Text, txtCod_TipAne, txtCod_Anexo, inpFec_Inicio.Text, inpFec_Final.Text)
Set gridex1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

Exit Sub
Resume
errores:
    errores err.Number
End Sub

Private Sub inpFec_Emi_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub





Private Sub txtDes_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 2, Me)
End Sub

Private Sub txtCod_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 1, Me)
End Sub

Private Sub txtNum_Parte_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub CambioEstado()
Dim strSQL As String
On Error GoTo errores

strSQL = "CN_VENTAS_CAMBIA_ESTADO_FACTURAS_DIFERIDAS '$' ,'$'"
strSQL = VBsprintf(strSQL, gridex1.Value(gridex1.Columns("NUM_CORRE").Index), vusu)

ExecuteCommandSQL cCONNECT, strSQL

Mensaje kMESSAGE_INF_PROCESS_SATISFACTO

Buscar
Exit Sub

Resume
errores:
    errores err.Number
End Sub


Private Sub txtRuc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    BUSCARUC 1
  End If
End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  gridex1.ClearFields
  If KeyAscii = vbKeyReturn Then
      If Trim(txtCod_TipAne.Text) <> "" Then
          Call BUSCA_TIPO_ANEXO(1, 1)
      Else
          Call BUSCA_TIPO_ANEXO(2, 1)
      End If
  End If
End Sub

Private Sub txtDes_Anexo_KeyPress(KeyAscii As Integer)
    gridex1.ClearFields
    If KeyAscii = vbKeyReturn Then
        If Trim(txtDes_Anexo.Text) <> "" Then
            If Len(Trim(txtDes_Anexo)) > 2 Then
                Call BUSCA_ANEXO(2, 1)
            Else
                Aviso "Debe ingresar al menos 3 caracteres del Nombre requerido", 1
                Exit Sub
            End If
        Else
            Aviso "Debe ingresar al menos 3 caracteres del Nombre requerido", 1
            Exit Sub
        End If
    End If
End Sub

Private Sub BUSCARUC(opcion As Integer)

On Error GoTo Fin
Dim strSQL As String
Dim oTipo As New frmBusqGeneral

    strSQL = "SELECT num_ruc as Ruc,Des_Anexo Descripcion FROM CN_AnexosContables "
    txtRuc = Trim(txtRuc)
    
    strSQL = strSQL & " where num_ruc like '%" & txtRuc & "%' and Cod_TipAnex ='C'"
    
    txtRuc = ""
        
    Set oTipo.oParent = Me
    
    oTipo.SQuery = strSQL
    oTipo.CARGAR_DATOS
    oTipo.DGridLista.Columns(1).Width = 4350.047
    oTipo.Show 1
    If codigo <> "" Then
      txtRuc = Trim(codigo)
      txtDes_Anexo = Trim(Descripcion)
      
      strSQL = "SELECT Cod_TipAnEx FROM CN_AnexosContables WHERE num_ruc = '" & txtRuc.Text & "' and Cod_TipAnex ='C'"
      txtCod_TipAne.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
      strSQL = "SELECT Cod_Anxo FROM CN_AnexosContables WHERE num_ruc = '" & txtRuc.Text & "' and Cod_TipAnex ='C'"
      txtCod_Anexo.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
      
      SendKeys "{TAB}"
    End If
    Set oTipo = Nothing
    
Exit Sub
Resume
Fin:
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda (" & opcion & ")"
End Sub


Sub BUSCA_ANEXO(Tipo As Integer, Ubic As Integer)

Dim iLen As Integer
    Select Case Tipo
        Case 1:
                If Ubic = 1 Then
                    strSQL = "SELECT MIN(DATALENGTH(COD_ANXO)) FROM CN_AnexosContables"
                    iLen = Trim(DevuelveCampo(strSQL, cCONNECT))
                    
                    txtCod_Anexo.Text = Right(Repl("0", iLen) & txtCod_Anexo, iLen)
                    
                     
                     strSQL = "SELECT Des_Anexo FROM CN_AnexosContables WHERE Cod_TipAnEX = '" & txtCod_TipAne.Text & "' AND Cod_Anxo = '" & txtCod_Anexo.Text & "'"
                     txtDes_Anexo.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                     SendKeys "{TAB}"
                     
                     Exit Sub
                     
                Else
                End If
        Case 2:
        
                Dim oTipo As New frmBusqGeneral
                Dim RS As Object
                Set RS = CreateObject("ADODB.Recordset")
                Set oTipo.oParent = Me
                If Ubic = 1 Then
                    oTipo.SQuery = "SELECT Cod_Anxo as Código, Des_Anexo as Descripción FROM CN_AnexosContables WHERE Cod_TipAnEX = '" & txtCod_TipAne.Text & "' AND Des_Anexo like '%" & Trim(txtDes_Anexo.Text) & "%'"
                Else
                End If
                oTipo.CARGAR_DATOS
                oTipo.Top = txtDes_Anexo.Top + txtDes_Anexo.Height
                oTipo.Left = txtDes_Anexo.Left
                oTipo.DGridLista.Columns(1).Width = 1000
                oTipo.Show 1
                If codigo <> "" Then
                    If Ubic = 1 Then
                        txtCod_Anexo.Text = Trim(codigo)
                        txtDes_Anexo.Text = Trim(Descripcion)
                        strSQL = "SELECT num_ruc FROM CN_AnexosContables WHERE Cod_TipAnEX = '" & txtCod_TipAne.Text & "' AND Cod_Anxo = '" & txtCod_Anexo.Text & "'"
                        txtRuc = Trim(DevuelveCampo(strSQL, cCONNECT))

                        SendKeys "{TAB}"
                    Else
                    End If
                End If
                Set oTipo = Nothing
                Set RS = Nothing
                
    End Select
    
End Sub



Sub BUSCA_TIPO_ANEXO(Tipo As Integer, Ubic As Integer)
    Select Case Tipo
        Case 1:
                If Ubic = 1 Then
                    strSQL = "SELECT DES_TIPANEX FROM CN_TipoAnexoContable WHERE COD_TIPANEX = '" & txtCod_TipAne.Text & "'"
                    txtCod_Anexo.SetFocus
                Else
                End If
        Case 2:
                Dim oTipo As New frmBusqGeneral
                Dim RS As Object
                Set RS = CreateObject("ADODB.Recordset")
                Set oTipo.oParent = Me
                If Ubic = 1 Then
                    oTipo.SQuery = "SELECT COD_TIPANEX as Código, DES_TIPANEX as Descripción FROM CN_TipoAnexoContable "
                Else
                End If
                oTipo.CARGAR_DATOS
                oTipo.Show 1
                If codigo <> "" Then
                    If Ubic = 1 Then
                        txtCod_TipAne.Text = Trim(codigo)
                        txtCod_Anexo.SetFocus
                    Else
                    End If
                End If
                Set oTipo = Nothing
                Set RS = Nothing
                
    End Select
End Sub

