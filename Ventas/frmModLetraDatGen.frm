VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmModLetraDatGen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificación de Datos Generales"
   ClientHeight    =   4335
   ClientLeft      =   1605
   ClientTop       =   1875
   ClientWidth     =   6570
   Icon            =   "frmModLetraDatGen.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6570
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1905
      TabIndex        =   10
      Top             =   3720
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmModLetraDatGen.frx":030A
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   3720
      Left            =   0
      TabIndex        =   11
      Top             =   -45
      Width           =   6480
      Begin VB.TextBox txtTercero_NomAnexo 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2970
         MaxLength       =   30
         TabIndex        =   9
         Top             =   3240
         Width           =   3225
      End
      Begin VB.TextBox txtTercero_Des_TipAnex 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4050
         MaxLength       =   11
         TabIndex        =   28
         Top             =   3240
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtTercero_NumRuc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   870
         MaxLength       =   11
         TabIndex        =   8
         Top             =   3240
         Width           =   1545
      End
      Begin VB.TextBox txtTercero_CodTipAnexo 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2565
         MaxLength       =   4
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "C"
         Top             =   3240
         Width           =   360
      End
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   870
         MaxLength       =   11
         TabIndex        =   5
         Top             =   2160
         Width           =   1545
      End
      Begin VB.TextBox TxtGlosa 
         Height          =   585
         Left            =   870
         TabIndex        =   7
         Top             =   2610
         Width           =   5325
      End
      Begin VB.TextBox txtDes_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2925
         MaxLength       =   30
         TabIndex        =   6
         Top             =   2160
         Width           =   3300
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   22
         Top             =   2160
         Width           =   360
      End
      Begin VB.TextBox TxtCod_Banco 
         Height          =   315
         Left            =   1590
         TabIndex        =   0
         Top             =   975
         Width           =   735
      End
      Begin VB.TextBox TxtNom_Banco 
         Height          =   315
         Left            =   2340
         TabIndex        =   1
         Top             =   975
         Width           =   3855
      End
      Begin VB.TextBox txtNumLetraBanco 
         Height          =   315
         Left            =   1590
         TabIndex        =   2
         Top             =   1320
         Width           =   1410
      End
      Begin VB.TextBox TxtMonto 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   270
         Left            =   4635
         TabIndex        =   17
         Top             =   270
         Width           =   1680
      End
      Begin VB.TextBox TxtNumero 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   255
         Width           =   1695
      End
      Begin NumBoxProject.NumBox inpFec_Venc 
         Height          =   315
         Left            =   4830
         TabIndex        =   4
         Top             =   1695
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
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
      Begin NumBoxProject.NumBox inpFec_Emi 
         Height          =   315
         Left            =   4830
         TabIndex        =   3
         Top             =   1350
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
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
      Begin VB.TextBox txtDes_TipAnex 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3360
         MaxLength       =   11
         TabIndex        =   23
         Top             =   2160
         Visible         =   0   'False
         Width           =   1545
      End
      Begin NumBoxProject.NumBox txtFecha_Banco_Desc 
         Height          =   315
         Left            =   1605
         TabIndex        =   29
         Top             =   1680
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Banco Dec. :"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   1710
         Width           =   1440
      End
      Begin VB.Label Label11 
         Caption         =   "Ruc Tercero"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   3135
         Width           =   615
      End
      Begin VB.Label Label28 
         Caption         =   "R.U.C."
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2175
         Width           =   615
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Aval :"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1965
         Width           =   405
      End
      Begin VB.Label Label16 
         Caption         =   "Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   990
         Width           =   885
      End
      Begin VB.Label Label10 
         Caption         =   "Num. Letra Banco :"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1350
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Glosa :"
         Height          =   255
         Left            =   150
         TabIndex        =   19
         Top             =   2610
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Emision :"
         Height          =   225
         Left            =   3600
         TabIndex        =   18
         Top             =   1395
         Width           =   1170
      End
      Begin VB.Label LblSimbolo 
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   270
         Left            =   4320
         TabIndex        =   16
         Top             =   270
         Width           =   270
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   195
         X2              =   6255
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Venc. :"
         Height          =   255
         Left            =   3585
         TabIndex        =   15
         Top             =   1725
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Monto :"
         Height          =   255
         Left            =   3585
         TabIndex        =   13
         Top             =   270
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Letra Nº :"
         Height          =   225
         Left            =   180
         TabIndex        =   12
         Top             =   255
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmModLetraDatGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mvarNum_Corr As String
Public codigo As String, Descripcion As String, TipoAdd As String, strCod_Anxo As String, strTercero_Cod_Anxo As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

On Error GoTo hand

    Select Case ActionName
        Case "ACEPTAR"
            SALVAR_DATOS
            Unload Me
        Case "CANCELAR"
            Unload Me
    End Select
    
    Exit Sub
hand:
errores Err.Number

End Sub

Sub SALVAR_DATOS()

Dim SQL As String

If Trim(inpFec_Emi.Text) = "" Then
    Aviso "Fecha Emisión es obligatoria", 1
    Exit Sub
End If

If Trim(inpFec_Venc.Text) = "" Then
    Aviso "Fecha Vencimieno es obligatoria", 1
    Exit Sub
End If

SQL = "Exec Ventas_Up_Man_Letras_Datos_Generales '" & mvarNum_Corr & "','" & TxtCod_Banco & "','" & txtNumLetraBanco & "','" & inpFec_Emi.Text & "','" & inpFec_Venc.Text & "','" & IIf(txtDes_TipAne = "", "", strCod_Anxo) & "','" & txtCod_TipAne & "','" & TxtGlosa.Text & "','" & txtTercero_CodTipAnexo & "','" & strTercero_Cod_Anxo & "','" & txtFecha_Banco_Desc.Text & "'"
ExecuteCommandSQL cCONNECT, SQL
  
End Sub

Private Sub inpFec_canje_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtCod_Banco.SetFocus
End Sub

Private Sub inpFec_Emi_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub inpFec_Venc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Label13_Click()

End Sub

Private Sub TxtCod_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("cod_banco", "nom_banco", "tg_banco where flg_operativo ='*' and ", TxtCod_Banco, TxtNom_Banco, 1, Me)
End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAnex, 1, Me)
End Sub
Private Sub txtDes_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 2, Me)
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then FunctButt1.SetFocus
End Sub

Private Sub TxtInteres_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtCod_Banco.SetFocus
End Sub

Private Sub TxtNom_Banco_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then If KeyAscii = 13 Then Call Busca_Opcion("cod_banco", "nom_banco", "tg_banco where flg_operativo ='*' and ", TxtCod_Banco, TxtNom_Banco, 2, Me)
End Sub

Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 1, Me)
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtNumLetraBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If inpFec_Venc.Enabled Then
            SendKeys "{TAB}"
        Else
            TxtGlosa.SetFocus
        End If
    End If
End Sub

Sub Busca_Banco(tipo As String)
Set frmBusqGeneral3.oParent = Me
If tipo = "1" Then
    frmBusqGeneral3.SQuery = "select cod_banco as Codigo ,nom_banco as Descripcion from tg_banco where cod_banco like '%" & Trim(TxtCod_Banco.Text) & "%' and flg_operativo ='*' order by 1"
Else
    frmBusqGeneral3.SQuery = "select cod_banco as Codigo ,nom_banco as Descripcion from tg_banco where nom_banco like '%" & Trim(TxtNom_Banco.Text) & "%' and flg_operativo ='*' order by 2"
End If
frmBusqGeneral3.Cargar_Datos
frmBusqGeneral3.gexLista.Columns("Descripcion").Width = "3000"

frmBusqGeneral3.Show 1
If codigo <> "" Then
    TxtCod_Banco.Text = codigo
    TxtNom_Banco.Text = Descripcion
    txtNumLetraBanco.SetFocus
Else
    TxtCod_Banco.Text = ""
    TxtNom_Banco.Text = ""
End If
Unload frmBusqGeneral3
Set frmBusqGeneral3 = Nothing
codigo = ""
Descripcion = ""
End Sub

Private Sub txtTercero_CodTipAnexo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtTercero_CodTipAnexo, txtTercero_Des_TipAnex, 1, Me)
End Sub

Private Sub txtTercero_NomAnexo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo_Tercero("Num_Ruc", "Des_Anexo", txtTercero_CodTipAnexo, txtTercero_NumRuc, txtTercero_NomAnexo, 2, Me)
End Sub

Private Sub txtTercero_NumRuc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo_Tercero("Num_Ruc", "Des_Anexo", txtTercero_CodTipAnexo, txtTercero_NumRuc, txtTercero_NomAnexo, 1, Me)
End Sub

Sub Busca_Opcion_Anexo_Tercero(strCampo1 As String, strCampo2 As String, StrTabla As String, txtCod As TextBox, txtDes As TextBox, Opcion As Integer, frmME As Form)

On Error GoTo Fin

Dim rstAux As ADODB.Recordset, strSQL As String
    strSQL = "select Cod_Anxo as Cod,Des_Anexo as Nombre,Num_Ruc as Ruc from cn_anexoscontables where cod_tipanex = '" & StrTabla & "' and "
    
    'StrSql = "Select " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & StrTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    Select Case Opcion
    Case 1: strSQL = strSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: strSQL = strSQL & strCampo2 & " like '%" & txtDes & "%'"
    End Select
    txtCod = ""
    txtDes = ""
    frmME.strTercero_Cod_Anxo = ""
    With frmBusqGeneral
        Set .oParent = frmME
        .SQuery = strSQL
        .Cargar_Datos
        
        codigo = ""
        .DGridLista.Columns("Cod").Visible = False
        .DGridLista.Columns("Nombre").Width = 4575
        .DGridLista.Columns("RUC").Width = 1695
        Set rstAux = .DGridLista.ADORecordset
        
        If rstAux.RecordCount > 1 Then
          .Show vbModal
        Else
          frmME.codigo = ".."
        End If
        If frmME.codigo <> "" And rstAux.RecordCount > 0 Then
            frmME.strTercero_Cod_Anxo = Trim(rstAux!Cod)
            txtDes = Trim(rstAux!Nombre)
            txtCod = Trim(rstAux!Ruc)
            Select Case Opcion
            Case 1: SendKeys "{TAB}"
            Case 2: SendKeys "{TAB}"
            End Select
        Else
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub


