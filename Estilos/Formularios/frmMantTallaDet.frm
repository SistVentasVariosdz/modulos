VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMantTallaDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tallas"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Sizes"
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   105
      TabIndex        =   8
      Top             =   4935
      Width           =   1965
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmMantTallaDet.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "frmMantTallaDet.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "frmMantTallaDet.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmMantTallaDet.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
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
      Height          =   3255
      Left            =   150
      TabIndex        =   7
      Tag             =   "List"
      Top             =   60
      Width           =   5445
      Begin GridEX20.GridEX DGridLista 
         Height          =   3465
         Left            =   90
         TabIndex        =   16
         Top             =   240
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   6112
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmMantTallaDet.frx":05C8
         Column(2)       =   "frmMantTallaDet.frx":0690
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmMantTallaDet.frx":0734
         FormatStyle(2)  =   "frmMantTallaDet.frx":086C
         FormatStyle(3)  =   "frmMantTallaDet.frx":091C
         FormatStyle(4)  =   "frmMantTallaDet.frx":09D0
         FormatStyle(5)  =   "frmMantTallaDet.frx":0AA8
         FormatStyle(6)  =   "frmMantTallaDet.frx":0B60
         ImageCount      =   0
         PrinterProperties=   "frmMantTallaDet.frx":0C40
      End
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
      Height          =   1545
      Left            =   150
      TabIndex        =   0
      Tag             =   "Detail"
      Top             =   3390
      Width           =   5445
      Begin VB.TextBox txtIdGrutal 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
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
         Left            =   1020
         MaxLength       =   3
         TabIndex        =   1
         Top             =   360
         Width           =   780
      End
      Begin VB.TextBox txtIdTalla 
         Enabled         =   0   'False
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
         Left            =   1020
         TabIndex        =   3
         Top             =   735
         Width           =   2025
      End
      Begin VB.TextBox txtIndice 
         Enabled         =   0   'False
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
         Left            =   1020
         TabIndex        =   4
         Top             =   1095
         Width           =   795
      End
      Begin VB.CommandButton cmdBuscaTalla 
         Caption         =   "..."
         Height          =   330
         Left            =   3060
         TabIndex        =   14
         Tag             =   "..."
         Top             =   720
         Width           =   285
      End
      Begin VB.TextBox txtDesGrutal 
         Enabled         =   0   'False
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
         Left            =   1815
         MaxLength       =   30
         TabIndex        =   2
         Top             =   360
         Width           =   3480
      End
      Begin VB.Label Label1 
         Caption         =   "Orden:"
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
         Left            =   255
         TabIndex        =   15
         Tag             =   "Order:"
         Top             =   1140
         Width           =   750
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
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
         Left            =   255
         TabIndex        =   6
         Tag             =   "Group:"
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Talla :"
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
         Left            =   255
         TabIndex        =   5
         Tag             =   "Size :"
         Top             =   795
         Width           =   750
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2055
      TabIndex        =   13
      Top             =   5010
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantTallaDet.frx":0E18
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantTallaDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs_Carga As New ADODB.Recordset
Public oParent As Object
Public Codigo As String
Public varCod_grutal As String
Dim vAccion As String
Dim vrow As Variant

Private Sub cmdBuscaTalla_Click()
Dim oTalla As New frmBusqGeneral
Set oTalla.oParent = Me
oTalla.sQuery = "SELECT cod_talla FROM TG_Talla"
oTalla.Cargar_Datos
oTalla.Show 1
Set oTalla = Nothing
If Codigo <> "" Then
    txtIdTalla.Text = Codigo
    Codigo = ""
End If
End Sub
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
    
'    B_sql = "SELECT * " & _
'    "FROM  ES_TallasDet " & _
'    "WHERE cod_grutal ='" & txtIdGrutal.Text & "' " & _
'    "AND   cod_talla  ='" & txtIdTalla.Text & "' "
'    Rs_DATOS.LockType = adLockOptimistic
'    Rs_DATOS.Open B_sql, B_conexion
'    If Rs_DATOS.EOF Then
'       Rs_DATOS.AddNew
'       Rs_DATOS!cod_grutal = txtIdGrutal.Text
'       Rs_DATOS!cod_talla = txtIdTalla.Text
'    End If
'    Rs_DATOS!num_indice = txtIndice.Text
'    Rs_DATOS.Update
'    Rs_DATOS.Close
    StrSQL = "UP_MAN_TALLASDET '" & vAccion & "','" & txtIdGrutal.Text & "','" & txtIdTalla.Text & "'," & Val(txtIndice.Text)
    ExecuteSQL cCONNECT, StrSQL

    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
    Informa "", amensaje
Exit Sub
Salvar_DatosErr:
ErrorHandler Err, "Salvar_Datos"
End Sub
Sub ELIMINAR_DATOS()
On Error GoTo Eliminar_DatosErr
If txtIdGrutal.Text <> "" Then
    'B_db.Execute ("DELETE FROM ES_TallasDet WHERE cod_grutal ='" & txtIdGrutal.Text & "' AND cod_talla = '" & txtIdTalla.Text & "'")
    B_db.Execute "UP_MAN_ES_TallasDet '" & txtIdGrutal & "','" & txtIdTalla & "',0,'D'"
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje
End If
LIMPIAR_DATOS
RECARGAR_DATOS
Exit Sub
Eliminar_DatosErr:
    ErrorHandler Err, "Eliminar_Datos"
End Sub
Sub LIMPIAR_DATOS()
  txtIdTalla.Text = ""
  txtIndice.Text = ""
  cmdBuscaTalla.Enabled = True
  txtIndice.Enabled = True
  cmdBuscaTalla.SetFocus
End Sub
Private Sub DGridLista_Click()
If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
    txtIdTalla.Text = Rs_Carga!cod_talla
    txtIndice.Text = Rs_Carga!num_indice
    DESHABILITA_DATOS
End If
End Sub
Sub DESHABILITA_DATOS()
txtIndice.Enabled = False
cmdBuscaTalla.Enabled = False
End Sub
Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub

Private Sub DGridLista_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
    txtIdTalla.Text = Rs_Carga!cod_talla
    txtIndice.Text = Rs_Carga!num_indice
    DESHABILITA_DATOS
End If
End Sub

Private Sub Form_Load()
Call FormSet(Me)
Realiza_Conexion
'FormateaGrid Me.DGridLista
B_conexion = cCONNECT
'DGridLista.Columns(0).DataField = "cod_talla"
'DGridLista.Columns(0).Caption = "Size"
'DGridLista.Columns(1).Caption = "Order"
'DGridLista.Columns(1).DataField = "num_indice"
Me.MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
vAccion = ""
End Sub
Public Sub Cargar_Datos()
On Error GoTo Cargar_DatosErr
Dim sIdTalla  As Long

vrow = DGridLista.Row
sIdTalla = Val(txtIdTalla.Text)

B_sql = "SELECT * FROM ES_TallasDet WHERE cod_grutal = '" & txtIdGrutal.Text & "' ORDER BY num_indice"
Rs_Carga.ActiveConnection = B_conexion
Rs_Carga.CursorType = adOpenStatic
Rs_Carga.Open B_sql
'Set DGridLista.DataSource = Rs_Carga
Set DGridLista.ADORecordset = Rs_Carga

If Not Rs_Carga.EOF Then
    txtIdTalla.Text = Rs_Carga!cod_talla
    txtIndice.Text = Rs_Carga!num_indice
    HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
Else
    HabilitaMant Me.MantFunc1, "ADICIONAR"
End If
DESHABILITA_DATOS
DGridLista.Columns(1).Width = 1000
DGridLista.Columns(2).Width = 1000
DGridLista.Columns(3).Width = 800

DGridLista.Columns(1).Caption = "Grupo"
DGridLista.Columns(2).Caption = "Talla"
DGridLista.Columns(3).Caption = "Indice"

If vAccion = "I" Then
    DGridLista.Find 2, jgexEqual, sIdTalla
Else
    DGridLista.Row = vrow
End If

Exit Sub
Cargar_DatosErr:
    ErrorHandler Err, "Cargar_Datos"
End Sub
Sub RECARGAR_DATOS()

Rs_Carga.Close
Cargar_Datos
End Sub
Sub BUSCA_GRUTALDET()
Dim Rs_busca As New ADODB.Recordset
On Error GoTo Busca_GruTalDetErr
B_sql = "Select * from ES_TallasDet WHERE " & _
"cod_grutal = '" & txtIdGrutal.Text & "' AND " & _
"cod_talla  = '" & txtIdTalla.Text & "'"
Rs_busca.ActiveConnection = cCONNECT
Rs_busca.CursorType = adOpenStatic
Rs_busca.Open B_sql
If Not Rs_busca.EOF Then
    txtIndice.Text = Rs_busca!num_indice
    DESHABILITA_DATOS
    HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    DGridLista.Enabled = True
End If
Rs_busca.Close
Set Rs_busca = Nothing
Exit Sub
Busca_GruTalDetErr:
    ErrorHandler Err, "Busca_GruTal"
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set Rs_Carga = Nothing
End Sub
Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "ADICIONAR"
        Load frmMantTallaDetalle
        frmMantTallaDetalle.Caption = "Seleccion de Tallas del Grupo:" & txtIdGrutal.Text & " - " & txtDesGrutal.Text
        frmMantTallaDetalle.varCod_grutal = txtIdGrutal.Text
        frmMantTallaDetalle.CARGA_TALLAS
        vAccion = "I"
        frmMantTallaDetalle.Show 1
        RECARGAR_DATOS
'        LIMPIAR_DATOS
'        cmdBuscaTalla.Enabled = True
'        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
'        DGridLista.Enabled = False
    Case "MODIFICAR"
        txtIndice.Enabled = True
        txtIndice.SetFocus
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        DGridLista.Enabled = False
        vAccion = "U"
    Case "ELIMINAR"
        'ELIMINAR_DATOS
        Load frmMantTallaDetalle
        frmMantTallaDetalle.Caption = "Seleccion de Tallas del Grupo:" & txtIdGrutal.Text & " - " & txtDesGrutal.Text
        frmMantTallaDetalle.varCod_grutal = txtIdGrutal.Text
        frmMantTallaDetalle.CARGA_TALLAS
        vAccion = "D"
        frmMantTallaDetalle.Show 1
        RECARGAR_DATOS
    Case "GRABAR"
        If VALIDA_DATOS Then
            SALVAR_DATOS
            RECARGAR_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
            cmdBuscaTalla.Enabled = False
            vAccion = ""
        End If
    Case "DESHACER"
        LIMPIAR_DATOS
        RECARGAR_DATOS
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        DGridLista.Enabled = True
        cmdBuscaTalla.Enabled = False
        vAccion = ""
    Case "SALIR"
        Unload Me
End Select
End Sub
Private Sub txtIdgrutal_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub
Private Sub txtIdTalla_LostFocus()
If txtIdGrutal.Text <> "" Then
    BUSCA_GRUTALDET
End If
End Sub
Private Sub txtIndice_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub
Function VALIDA_DATOS() As Boolean
Dim aMess(4)
Dim amensaje As clsMessages
Set amensaje = New clsMessages
VALIDA_DATOS = True
If Len(txtIndice) = 0 Then
   MsgBox "Ingrese el Orden", vbInformation, Me.Caption
   VALIDA_DATOS = False
End If
If Len(txtIdTalla) = 0 Then
   MsgBox "Ingrese la Talla", vbInformation, Me.Caption
   VALIDA_DATOS = False
End If
If Not VALIDA_DATOS Then
    LoadMessage aMess, amensaje.Codigo
    amensaje.ShowMesage (iLanguage)
End If
End Function
