VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmAjusteLetra 
   Caption         =   "Ajuste de la Letra"
   ClientHeight    =   4725
   ClientLeft      =   2265
   ClientTop       =   3810
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   8835
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   8535
      Begin VB.CheckBox chkRetencion 
         Caption         =   "No Aplica Sobre Importe Total"
         Height          =   195
         Left            =   2880
         TabIndex        =   0
         Top             =   650
         Width           =   2535
      End
      Begin NumBoxProject.NumBox txtImpTotalN 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         TypeVal         =   2
         Mask            =   "9,999,999,999.99"
         Formato         =   "#,###,###,###.##"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.00"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   2
      End
      Begin NumBoxProject.NumBox txtImpNetoN 
         Height          =   315
         Left            =   6840
         TabIndex        =   2
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         TypeVal         =   2
         Mask            =   "9,999,999,999.99"
         Formato         =   "#,###,###,###.##"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.00"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   2
      End
      Begin NumBoxProject.NumBox txtImpRetencionN 
         Height          =   315
         Left            =   4320
         TabIndex        =   1
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         TypeVal         =   2
         Mask            =   "9,999,999,999.99"
         Formato         =   "#,###,###,###.##"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.00"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   2
      End
      Begin VB.Label Label6 
         Caption         =   "Importe Neto :"
         Height          =   255
         Left            =   5760
         TabIndex        =   16
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Importe Retencion :"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Importe Total :"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   8535
      Begin VB.TextBox txtImpTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   240
         Width           =   1245
      End
      Begin VB.TextBox txtImpRetencion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   4320
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   240
         Width           =   1245
      End
      Begin VB.TextBox txtImpNeto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   6840
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label4 
         Caption         =   "Importe Total :"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Retencion :"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Importe Neto :"
         Height          =   255
         Left            =   5760
         TabIndex        =   10
         Top             =   270
         Width           =   1035
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3120
      TabIndex        =   4
      Top             =   4080
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmAjusteLetra.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX gexLetra 
      Height          =   1800
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   3175
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GridLineStyle   =   2
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      BorderStyle     =   2
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmAjusteLetra.frx":0096
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmAjusteLetra.frx":03B0
      Column(2)       =   "frmAjusteLetra.frx":0478
      FormatStylesCount=   9
      FormatStyle(1)  =   "frmAjusteLetra.frx":051C
      FormatStyle(2)  =   "frmAjusteLetra.frx":0654
      FormatStyle(3)  =   "frmAjusteLetra.frx":0704
      FormatStyle(4)  =   "frmAjusteLetra.frx":07B8
      FormatStyle(5)  =   "frmAjusteLetra.frx":0890
      FormatStyle(6)  =   "frmAjusteLetra.frx":0948
      FormatStyle(7)  =   "frmAjusteLetra.frx":0A28
      FormatStyle(8)  =   "frmAjusteLetra.frx":0E34
      FormatStyle(9)  =   "frmAjusteLetra.frx":1244
      ImageCount      =   1
      ImagePicture(1) =   "frmAjusteLetra.frx":13CC
      PrinterProperties=   "frmAjusteLetra.frx":16E6
   End
   Begin VB.Label Label7 
      Caption         =   "Letra a Afectar Ajuste"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   1800
      Width           =   1635
   End
End
Attribute VB_Name = "frmAjusteLetra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NumCorre, strSQL As String
Public sFlg_Retencion_IGV As String

Dim Factor As Double, Corre_Afecto As String

Private Sub chkRetencion_Click()
  txtImpRetencionN.SetFocus
End Sub

Private Sub Form_Load()
  Factor = DevuelveCampo("select 1- dbo.CN_Obtiene_Factor_Retencion()", cCONNECT)
  CARGA_GRID
End Sub

Sub CARGA_GRID()

On Error GoTo hand
    Set gexLetra.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
Exit Sub

hand:
ErrorHandler Err, "CARGA_GRID"

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand
Select Case ActionName
Case "ACEPTAR"
    Corre_Afecto = IIf(gexLetra.RowCount = 0, "", gexLetra.Value(gexLetra.Columns("Correlativo").Index))
    If MsgBox("Esta seguro de hacer un ajuste a esta Letra", vbYesNo, "AVISO") = vbYes Then
      Call ExecuteSQL(cCONNECT, "UP_CAMBIA_IMPORTE_LETRAS '" & NumCorre & "'," & txtImpTotalN.Text & "," & txtImpRetencionN.Text & "," & txtImpNetoN.Text & ",'" & Corre_Afecto & "','" & vusu & "'")
      
'      GeneraInterfase_Cn_Docum_New NumCorre, "", "U"
'      If RTrim(Corre_Afecto) <> "" Then
'        GeneraInterfaseFINANZAS_Pagos_LETRAS Corre_Afecto, "", "", "U"
'      End If
'      GeneraInterfaseFINANZAS_Pagos_LETRAS NumCorre, "", "", "U"
      
      MsgBox "El Ajuste se ha llevado Satisfactoriamente", vbInformation, "AVISO"
      Unload Me
    End If
Case "CANCELAR"
  Unload Me
End Select
Exit Sub
hand:
ErrorHandler Err, "CARGA_GRID"
End Sub

Private Sub txtImpNetoN_Change()
 If txtImpNetoN.Text = "" Then txtImpNetoN = 0
 If chkRetencion = 0 Then
   If UCase(sFlg_Retencion_IGV) = "S" Then
     txtImpTotalN.Text = Format(CDbl(txtImpNetoN.Text) / Factor, "######.00")
     txtImpRetencionN.Text = CDbl(txtImpTotalN.Text) - CDbl(txtImpNetoN.Text)
  End If
 End If
End Sub

Private Sub txtImpRetencionN_Change()
  If chkRetencion = 0 Then txtImpTotalN.Text = CDbl(txtImpNetoN.Text) + CDbl(txtImpRetencionN.Text)
End Sub

