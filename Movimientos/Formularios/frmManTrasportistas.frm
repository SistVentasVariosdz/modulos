VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmManTransportistas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transportista"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "frmManTrasportistas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2040
      TabIndex        =   6
      Top             =   5160
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmManTrasportistas.frx":030A
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin VB.TextBox TxtPeso 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   5880
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox TxtLicencia 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox TxtRegTrans 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5880
      TabIndex        =   2
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox TxtNomConductor 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   4200
      Width           =   6015
   End
   Begin VB.TextBox TxtMarcaPlaca 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   1
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox TxtSecuencia 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      TabIndex        =   14
      Top             =   3480
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   7695
      Begin GridEX20.GridEX gexTrans 
         Height          =   2775
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4895
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmManTrasportistas.frx":046A
         Column(2)       =   "frmManTrasportistas.frx":0532
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmManTrasportistas.frx":05D6
         FormatStyle(2)  =   "frmManTrasportistas.frx":070E
         FormatStyle(3)  =   "frmManTrasportistas.frx":07BE
         FormatStyle(4)  =   "frmManTrasportistas.frx":0872
         FormatStyle(5)  =   "frmManTrasportistas.frx":094A
         FormatStyle(6)  =   "frmManTrasportistas.frx":0A02
         ImageCount      =   0
         PrinterProperties=   "frmManTrasportistas.frx":0AE2
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Kg"
      Height          =   255
      Left            =   7440
      TabIndex        =   15
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "Numero Licencia :"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Nombre Conductor :"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4310
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Peso Vehicular :"
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Reg. Transportista :"
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   3885
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Marca y Placa :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3870
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Secuencia :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
End
Attribute VB_Name = "frmManTransportistas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim Estado As String
Dim sTipo As String

Sub CARGA_GRID()
On Error GoTo hand

strSQL = "EXEC UP_SEL_TRANSPORTISTA"

Set Me.gexTrans.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
'ConfigurarGrid

Exit Sub
hand:
ErrorHandler err, "CARGA_GRID"
End Sub

Sub ConfigurarGrid()
        gexTrans.Columns("licencia").Visible = False
        gexTrans.Columns("Reg. Transportista").Visible = False
        
        gexTrans.Columns("Secuencia").Width = 1000
        gexTrans.Columns("conductor").Width = 2500
End Sub


Private Sub Form_Load()
    CARGA_GRID
End Sub

Sub CARGA_DATOS()
On Error GoTo hand
    TxtSecuencia.Text = gexTrans.Value(gexTrans.Columns("secuencia").Index)
    TxtMarcaPlaca.Text = Trim(gexTrans.Value(gexTrans.Columns("Marca y Placa").Index))
    TxtNomConductor.Text = Trim(gexTrans.Value(gexTrans.Columns("Conductor").Index))
    TxtPeso.Text = Trim(gexTrans.Value(gexTrans.Columns("peso").Index))
    TxtLicencia.Text = Trim(gexTrans.Value(gexTrans.Columns("licencia").Index))
    TxtRegTrans.Text = Trim(gexTrans.Value(gexTrans.Columns("Reg. Transportista").Index))
Exit Sub
hand:
ErrorHandler err, "CARGA_DATOS"
End Sub

Private Sub gexTrans_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    If gexTrans.RowCount = 0 Then Exit Sub
    CARGA_DATOS
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand
Select Case ActionName
    Case "ADICIONAR"
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Limpia
        HABILITA_CAMPOS True
        Estado = "NUEVO"
        sTipo = "I"
        strSQL = "select isnull(max(secuencia),0)+1 from lg_transportista"
        TxtSecuencia.Text = Format(DevuelveCampo(strSQL, cConnect), "000")
        TxtMarcaPlaca.SetFocus
    Case "MODIFICAR"
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Estado = "MODIFICAR"
        sTipo = "U"
        HABILITA_CAMPOS True
        TxtNomConductor.SetFocus
    Case "ELIMINAR"
        sTipo = "D"
        SALVAR_DATOS
        Limpia
        CARGA_GRID
        HABILITA_CAMPOS False
        sTipo = ""
    Case "GRABAR"
        If Trim(Me.TxtMarcaPlaca.Text) = "" Then MsgBox "Ingrese el numero de placa del auto", vbInformation: Exit Sub
'        If CmbTipItem = "" Then MsgBox "Seleccione un tipo de item", vbInformation: Exit Sub
        
        SALVAR_DATOS
        Limpia
        HABILITA_CAMPOS False
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        CARGA_GRID
        sTipo = ""
    Case "DESHACER"
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Limpia
        CARGA_GRID
        HABILITA_CAMPOS False
        sTipo = ""
    Case "SALIR"
        Unload Me
End Select

Exit Sub
hand:
ErrorHandler err, "MantFunc1_ActionClick"
End Sub

Sub HABILITA_CAMPOS(vEstado As Boolean)
    TxtMarcaPlaca.Enabled = vEstado
    TxtLicencia.Enabled = vEstado
    TxtRegTrans.Enabled = vEstado
    TxtNomConductor.Enabled = vEstado
    TxtPeso.Enabled = vEstado
End Sub

Sub Limpia()
    TxtMarcaPlaca.Text = ""
    TxtLicencia.Text = ""
    TxtRegTrans.Text = ""
    TxtNomConductor = ""
    TxtPeso = "0.00"
End Sub

Sub SALVAR_DATOS()
On Error GoTo hand
    strSQL = "EXEC UP_MAN_TRANSPORTISTA '" & sTipo & "','" & TxtSecuencia.Text & "','" & Trim(Me.TxtMarcaPlaca.Text) & "','" & Trim(Me.TxtRegTrans.Text) & "'," & Me.TxtPeso.Text & ",'" & Trim(Me.TxtNomConductor.Text) & "','" & Trim(Me.TxtLicencia.Text) & "'"
    
    Call ExecuteSQL(cConnect, strSQL)
    
Exit Sub
hand:
ErrorHandler err, "SALVAR_DATOS"
End Sub

Private Sub TxtLicencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.TxtPeso.SetFocus
End Sub

Private Sub TxtMarcaPlaca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.TxtRegTrans.SetFocus
End Sub

Private Sub TxtPeso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.MantFunc1.SetFocus
    Else
        Call SoloNumeros(TxtPeso, KeyAscii, True, 2)
    End If
End Sub

Private Sub TxtRegTrans_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.TxtNomConductor.SetFocus
End Sub


