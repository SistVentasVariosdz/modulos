VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "mantenimientos.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMantESTalla 
   Caption         =   "Grupo Tallas"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Size Group"
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   495
      Left            =   5610
      TabIndex        =   12
      Top             =   90
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      Custom          =   "0~0~DETALLE~Verdadero~Verdadero~&Detalle~0~0~1~~0~Falso~Falso~&Detalle~"
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Fradetalle 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   120
      TabIndex        =   9
      Tag             =   "Detail"
      Top             =   3930
      Width           =   5445
      Begin VB.TextBox txtDesGrutal 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Text            =   "50"
         Top             =   720
         Width           =   3585
      End
      Begin VB.TextBox txtIdGrutal 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   375
         TabIndex        =   11
         Tag             =   "Description"
         Top             =   795
         Width           =   945
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   375
         TabIndex        =   10
         Tag             =   "Code"
         Top             =   420
         Width           =   945
      End
   End
   Begin VB.Frame Fralista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   8
      Tag             =   "List"
      Top             =   30
      Width           =   5445
      Begin GridEX20.GridEX DGridLista 
         Height          =   3465
         Left            =   90
         TabIndex        =   13
         Top             =   240
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   6112
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmMantESTalla.frx":0000
         Column(2)       =   "frmMantESTalla.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmMantESTalla.frx":016C
         FormatStyle(2)  =   "frmMantESTalla.frx":02A4
         FormatStyle(3)  =   "frmMantESTalla.frx":0354
         FormatStyle(4)  =   "frmMantESTalla.frx":0408
         FormatStyle(5)  =   "frmMantESTalla.frx":04E0
         FormatStyle(6)  =   "frmMantESTalla.frx":0598
         ImageCount      =   0
         PrinterProperties=   "frmMantESTalla.frx":0678
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   75
      TabIndex        =   3
      Top             =   5175
      Width           =   1965
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmMantESTalla.frx":0850
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "frmMantESTalla.frx":09C2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "frmMantESTalla.frx":0B34
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmMantESTalla.frx":0CA6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2025
      TabIndex        =   0
      Top             =   5250
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantESTalla.frx":0E18
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   5970
      Top             =   5250
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmMantESTalla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs_Carga As New ADODB.Recordset
Dim sql1 As String
Dim sql2 As String
Dim sql3 As String

Dim Matriz As Variant
Dim Mayor As Integer

Dim Estado As String

Private Sub cmdFirst_Click()
If Not Rs_Carga.BOF Then
  Rs_Carga.MoveFirst
End If
End Sub
Private Sub cmdLast_Click()
If Not Rs_Carga.EOF Then
 Rs_Carga.MoveLast
End If
End Sub
Private Sub cmdNext_Click()
If Not Rs_Carga.EOF Then
 Rs_Carga.MoveNext
End If
End Sub
Private Sub cmdPrevious_Click()
If Not Rs_Carga.BOF Then
 Rs_Carga.MovePrevious
End If
End Sub
Sub SALVAR_DATOS()
On Error GoTo Salvar_DatosErr
    Dim Rs_DATOS As New ADODB.Recordset
    Dim rs_Det As New ADODB.Recordset
    Dim i As Integer
'    B_sql = "SELECT * " & _
'    "FROM  ES_Tallas " & _
'    "WHERE cod_grutal ='" & txtIdGrutal.Text & "'"
    

    Rs_DATOS.ActiveConnection = cCONNECT
    Rs_DATOS.CursorType = adOpenStatic
    Rs_DATOS.CursorLocation = adUseClient

    If Estado = "Nuevo" Then
        Rs_DATOS.Open "EXEC UP_MAN_ES_TALLA 'I','" & txtIdGrutal.Text & "','" & txtDesGrutal & "'"
        'Aqui ingresamos el detalle
        If Mayor > 0 Then
            rs_Det.ActiveConnection = cCONNECT
            rs_Det.CursorType = adOpenStatic
            rs_Det.CursorLocation = adUseClient

            For i = 0 To Mayor
                rs_Det.Open "EXEC UP_MAN_ES_TALLADET 'I','" & Trim(txtIdGrutal.Text) & "','" & Trim(Matriz(i)) & "'," & (i + 1) & ""

'                rs_Det!cod_grutal = Trim(txtIdGrutal.Text)
'                rs_Det!cod_talla = Trim(Matriz(i))
'                rs_Det!num_indice = (i + 1)
'                rs_Det.Update
            Next
        End If
    Else
        Rs_DATOS.Open "EXEC UP_MAN_ES_TALLA 'U','" & txtIdGrutal.Text & "','" & txtDesGrutal & "'"
    End If
    
    DESHABILITA_DATOS
    Cargar_Datos
    
'    If Rs_DATOS.EOF Then
'       Rs_DATOS.AddNew
'       Rs_DATOS!cod_grutal = txtIdGrutal.Text
'    End If
'    Rs_DATOS!des_grutal = txtDesGrutal.Text
'    Rs_DATOS.Update
'    Rs_DATOS.Close
'    Dim amensaje As New clsMessages
'    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
'    Informa "", amensaje
Exit Sub
Salvar_DatosErr:
ErrorHandler Err, "Salvar_Datos"
End Sub
Sub ELIMINAR_DATOS()
Dim Rs_DATOS As New ADODB.Recordset
On Error GoTo Eliminar_DatosErr
If txtIdGrutal.Text <> "" Then
    Rs_DATOS.ActiveConnection = cCONNECT
    Rs_DATOS.CursorType = adOpenStatic
    Rs_DATOS.CursorLocation = adUseClient

    Rs_DATOS.Open "EXEC UP_MAN_ES_TALLA 'D','" & txtIdGrutal.Text & "','" & txtDesGrutal & "'"

End If
LIMPIAR_DATOS
RECARGAR_DATOS
Exit Sub
Eliminar_DatosErr:
    ErrorHandler Err, "Eliminar_Datos"
End Sub
Sub LIMPIAR_DATOS()
  txtIdGrutal.Text = ""
  txtDesGrutal.Text = ""
  txtIdGrutal.Enabled = True
  txtDesGrutal.Enabled = True
  txtIdGrutal.SetFocus
End Sub

Sub DESHABILITA_DATOS()
txtIdGrutal.Enabled = False
txtDesGrutal.Enabled = False
End Sub
Sub HABILITA_DATOS()
txtIdGrutal.Enabled = True
txtDesGrutal.Enabled = True
End Sub


Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub

Private Sub DGridLista_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
If DGridLista.RowCount > 0 Then
    txtIdGrutal.Text = DGridLista.Value(DGridLista.Columns(1).Index)
    txtDesGrutal.Text = DGridLista.Value(DGridLista.Columns(2).Index)
    DESHABILITA_DATOS
End If
End Sub

Private Sub Form_Load()
Call FormSet(Me)
Realiza_Conexion
'FormateaGrid Me.DGridLista
B_conexion = cCONNECT
'Me.MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)

'DGridLista.Columns(0).DataField = "cod_grutal"
'DGridLista.Columns(1).DataField = "des_grutal"

Cargar_Datos
MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
Me.FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Sub Cargar_Datos()
On Error GoTo Cargar_DatosErr

B_sql = "SELECT * FROM ES_Tallas"
Set Rs_Carga = New ADODB.Recordset
Rs_Carga.ActiveConnection = B_conexion
Rs_Carga.CursorType = adOpenStatic
Rs_Carga.Open B_sql
'Set DGridLista.DataSource = Rs_Carga

Set DGridLista.ADORecordset = Rs_Carga

If Not Rs_Carga.EOF Then
    txtIdGrutal.Text = Rs_Carga!cod_grutal
    txtDesGrutal.Text = Rs_Carga!des_grutal
End If
DESHABILITA_DATOS
DGridLista.Columns(1).Width = 900
DGridLista.Columns(2).Width = 2500
DGridLista.Columns(1).Caption = "Código"
DGridLista.Columns(2).Caption = "Descripción"

Exit Sub
Cargar_DatosErr:
    ErrorHandler Err, "Cargar_Datos"
End Sub
Sub RECARGAR_DATOS()
Cargar_Datos
End Sub
Sub BUSCA_GRUTAL()
On Error GoTo Busca_GruTalErr
Dim Rs_busca As New ADODB.Recordset
B_sql = "Select * from ES_Tallas WHERE " & _
"cod_grutal = '" & txtIdGrutal.Text & "'"
Rs_busca.ActiveConnection = cCONNECT
Rs_busca.CursorType = adOpenStatic
Rs_busca.Open B_sql
If Not Rs_busca.EOF Then
    txtDesGrutal.Text = Rs_busca!des_grutal
    DESHABILITA_DATOS
    HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    DGridLista.Enabled = True
End If
Rs_busca.Close
Set Rs_busca = Nothing
Exit Sub
Busca_GruTalErr:
    ErrorHandler Err, "Busca_GruTal"
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set Rs_Carga = Nothing
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "DETALLE"
        Dim oMantDet As New frmMantTallaDet
        'Load oMantDet
        Set oMantDet.oParent = Me
        oMantDet.txtDesGrutal.Text = txtDesGrutal.Text
        oMantDet.txtIdGrutal.Text = txtIdGrutal.Text
        oMantDet.Cargar_Datos
        oMantDet.Show 1
        Set oMantDet = Nothing
End Select
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "ADICIONAR"
        Estado = "Nuevo"
        LIMPIAR_DATOS
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        DGridLista.Enabled = False
    Case "MODIFICAR"
        Estado = "Actualizar"
        txtIdGrutal.Enabled = False
        txtDesGrutal.Enabled = True
        txtDesGrutal.SetFocus
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        DGridLista.Enabled = False
    Case "ELIMINAR"
        Estado = "Eliminar"
        ELIMINAR_DATOS
    Case "GRABAR"
        If VALIDA_DATOS Then
            SALVAR_DATOS
            RECARGAR_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
        End If
    Case "DESHACER"
        Estado = "Deshacer"
        LIMPIAR_DATOS
        RECARGAR_DATOS
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        DGridLista.Enabled = True
    Case "SALIR"
        Unload Me
End Select
End Sub
Private Sub txtIdgrutal_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub
Private Sub txtIdGrutal_LostFocus()
If txtIdGrutal.Text <> "" Then
    BUSCA_GRUTAL
End If
End Sub
Private Sub txtdesgrutal_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub
Function VALIDA_DATOS() As Boolean
Dim aMess(4)
Dim amensaje As clsMessages
Dim i As Integer
Set amensaje = New clsMessages
VALIDA_DATOS = True
If Len(txtDesGrutal) = 0 Then
   MsgBox "Ingrese la descripcion de Grupo de Talla", vbInformation, Me.Caption
   VALIDA_DATOS = False
End If
If Len(txtIdGrutal) = 0 Then
   MsgBox "Ingrese la Codigo de Grupo de Talla", vbInformation, Me.Caption
   VALIDA_DATOS = False
End If

If Len(txtDesGrutal) > 0 Then
    Matriz = Split(txtDesGrutal.Text, ",")
            Mayor = UBound(Matriz)
            If Mayor > 0 Then
                For i = 0 To Mayor
                  If Trim(Matriz(i)) = "" Then
                    MsgBox "Se encontró un error en la Descripcion de la Talla, sirvase revisar los datos", vbInformation, "Tallas"
                    txtDesGrutal.SetFocus
                    VALIDA_DATOS = False
                    Exit Function
                  End If
                  If Len(Trim(Matriz(i))) > 10 Then
                    MsgBox "La longitud de las Tallas no puede exceder de 10 caracteres, sirvase revisar los datos", vbInformation, "Tallas"
                    txtDesGrutal.SetFocus
                    VALIDA_DATOS = False
                    Exit Function
                  End If
                Next i
            End If
End If

If Not VALIDA_DATOS Then
    LoadMessage aMess, amensaje.Codigo
    amensaje.ShowMesage (iLanguage)
End If
End Function
