VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmDetDatosTecnicos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos Tecnicos"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   Icon            =   "frmDetDatosTecnicos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Width           =   6795
      Begin VB.TextBox TxtTela 
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
         Height          =   315
         Left            =   3120
         TabIndex        =   34
         Top             =   200
         Width           =   3495
      End
      Begin VB.TextBox txtPartida 
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
         Height          =   315
         Left            =   720
         TabIndex        =   32
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Tela :"
         Height          =   255
         Left            =   2640
         TabIndex        =   33
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Partida :"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   615
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   570
      Left            =   2055
      TabIndex        =   15
      Top             =   6540
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   1005
      Custom          =   "0~0~ACEPTAR~True~True~&Aceptar~0~0~1~~0~False~False~&Aceptar~~1~0~CANCELAR~True~True~&Cancelar~1~0~3~~0~False~False~&Cancelar~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1300
      ControlHeigth   =   550
      ControlSeparator=   150
   End
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   6795
      Begin VB.TextBox TxtGramajePP 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TxtAnchoPP 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0"
         Top             =   975
         Width           =   1095
      End
      Begin VB.TextBox TxtReviradoPP 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "Adicionales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   120
         TabIndex        =   35
         Top             =   3315
         Width           =   6570
         Begin VB.TextBox txtKilosRechazados 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5295
            TabIndex        =   49
            Text            =   "0"
            Top             =   615
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox TxtElongacionAncho 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   4590
            TabIndex        =   46
            Text            =   "0"
            Top             =   1590
            Width           =   1110
         End
         Begin VB.CommandButton CmdNuevo_MotRechazo 
            Caption         =   "..."
            Height          =   330
            Left            =   5880
            TabIndex        =   45
            Top             =   1930
            Width           =   540
         End
         Begin VB.TextBox TxtDes_MotRechazo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2625
            TabIndex        =   44
            Top             =   1930
            Width           =   3105
         End
         Begin VB.TextBox TxtElongacionLargo 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   41
            Text            =   "0"
            Top             =   1590
            Width           =   1110
         End
         Begin VB.TextBox TxtCod_MotRechazo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   40
            Top             =   1930
            Width           =   900
         End
         Begin VB.TextBox txtGramaje_LavadoPP 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "0"
            Top             =   285
            Width           =   1110
         End
         Begin VB.ComboBox cboFlg_Status_Apro 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   630
            Width           =   1785
         End
         Begin VB.TextBox txtGramaje_Lavado 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   11
            Text            =   "0"
            Top             =   285
            Width           =   1110
         End
         Begin VB.TextBox txtObservaciones 
            Height          =   495
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   1020
            Width           =   4755
         End
         Begin VB.Label Label18 
            Caption         =   "Kgs Kilos Rechazados:"
            Height          =   255
            Left            =   3510
            TabIndex        =   48
            Top             =   675
            Width           =   1710
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Elongacion Ancho"
            Height          =   195
            Left            =   3150
            TabIndex        =   47
            Top             =   1575
            Width           =   1305
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Motivo de Rechazo"
            Height          =   195
            Left            =   210
            TabIndex        =   43
            Top             =   1995
            Width           =   1395
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Elongacion Largo"
            Height          =   195
            Left            =   210
            TabIndex        =   42
            Top             =   1575
            Width           =   1245
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Status"
            Height          =   195
            Left            =   150
            TabIndex        =   38
            Top             =   680
            Width           =   570
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Gramaje Lavado"
            Height          =   195
            Left            =   135
            TabIndex        =   37
            Top             =   270
            Width           =   1170
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   150
            TabIndex        =   36
            Top             =   1050
            Width           =   1110
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Encogimientos"
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
         TabIndex        =   24
         Top             =   1755
         Width           =   6570
         Begin VB.TextBox TxtEncogAPP 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "0"
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox TxtEncogLPP 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "0"
            Top             =   975
            Width           =   1095
         End
         Begin VB.ComboBox CboTipo 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox TxtEncLarPrg 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   5280
            TabIndex        =   28
            Text            =   "0"
            Top             =   1005
            Width           =   1095
         End
         Begin VB.TextBox TxtEncAncPrg 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   5280
            TabIndex        =   27
            Text            =   "0"
            Top             =   630
            Width           =   1095
         End
         Begin VB.TextBox TxtEncogA 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1680
            TabIndex        =   7
            Text            =   "0"
            Top             =   630
            Width           =   1095
         End
         Begin VB.TextBox TxtEncogL 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1680
            TabIndex        =   8
            Text            =   "0"
            Top             =   1005
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo :"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Encog. Ancho :"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   750
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Encog. Largo"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1125
            Width           =   1215
         End
      End
      Begin VB.TextBox TxtRevirado 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Text            =   "0"
         Top             =   1365
         Width           =   1095
      End
      Begin VB.TextBox TxtAnchoPrg 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   5400
         TabIndex        =   22
         Text            =   "0"
         Top             =   975
         Width           =   1095
      End
      Begin VB.TextBox TxtGramPrg 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   5400
         TabIndex        =   21
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TxtAncho 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Text            =   "0"
         Top             =   975
         Width           =   1095
      End
      Begin VB.TextBox TxtGramaje 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Planta Propia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3855
         TabIndex        =   39
         Top             =   170
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Revirado :"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1485
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Programado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Real"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2175
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Ancho (Mt) :"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1095
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Gramaje (Gr/Mt2) :"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmDetDatosTecnicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vCod_TipOrdTra As String
Public vCod_OrdTra As String
Public vCod_Tela As String
Public vDes_tela As String
Public vCod_Comb As String
Public vCod_Color As String
Public vPartida As String

Public Codigo As String
Public Descripcion As String

Private Sub cboFlg_Status_Apro_Click()
    If Right(cboFlg_Status_Apro.Text, 1) = "A" Then
       txtKilosRechazados.Text = 0
       txtKilosRechazados.Visible = False
    Else
       txtKilosRechazados.Text = 0
       txtKilosRechazados.Visible = True
    End If
End Sub

Private Sub CmdNuevo_MotRechazo_Click()
Load FrmNuevoMot_Rechazo
FrmNuevoMot_Rechazo.Show 1
End Sub

Private Sub Form_Load()
Dim Rs As ADODB.Recordset
Dim RsTelas As ADODB.Recordset
On Error GoTo hand

LlenaCombo CboTipo, "Select des_tipenc+space(100)+ cod_tipenc from tx_tipos_enc order by 1", cConnect

cboFlg_Status_Apro.Clear
cboFlg_Status_Apro.AddItem "Aprobado" & Space(100) & "A"
cboFlg_Status_Apro.AddItem "Desaprobado" & Space(100) & "D"
cboFlg_Status_Apro.ListIndex = 0


Set Rs = New ADODB.Recordset
Rs.ActiveConnection = cConnect
Rs.CursorLocation = adUseClient
Rs.CursorType = adOpenStatic

TxtPartida.Text = Trim(vPartida)
TxtTela = Trim(vCod_Tela) & "-" & Trim(vDes_tela)

Rs.Open "select * from tx_ordtra_telas where Cod_TipOrdTra='" & vCod_TipOrdTra & "' and Cod_OrdTra='" & vCod_OrdTra & "' and Cod_Tela='" & vCod_Tela & "' and Cod_comb='" & vCod_Comb & "' and Cod_Color='" & vCod_Color & "'"
If Rs.RecordCount Then
    TxtGramaje.Text = Rs("gramaje_acab")
    TxtAncho.Text = Rs("Ancho_acab")
    TxtEncogA.Text = Rs("Encog_Ancho")
    TxtEncogL.Text = Rs("Encog_Largo")
    TxtRevirado.Text = Rs("Revirado")
    
    TxtGramajePP.Text = Rs("gramaje_acab_PlantaPropia")
    TxtAnchoPP.Text = Rs("Ancho_acab_PlantaPropia")
    TxtEncogAPP.Text = Rs("Encog_Ancho_PlantaPropia")
    TxtEncogLPP.Text = Rs("Encog_Largo_PlantaPropia")
    TxtReviradoPP.Text = Rs("Revirado_PlantaPropia")
    If IsNull(Rs("Gramaje_Lavado_PlantaPropia").Value) Then
        Me.txtGramaje_LavadoPP.Text = ""
    Else
        Me.txtGramaje_LavadoPP.Text = Rs("Gramaje_Lavado_PlantaPropia").Value
    End If

    Call BuscaCombo(Rs("cod_Tipenc"), 1, CboTipo)
    
    'Esto es para los datos Adicionales
    If IsNull(Rs("Gramaje_Lavado").Value) Then
        Me.txtGramaje_Lavado.Text = ""
    Else
        Me.txtGramaje_Lavado.Text = Rs("Gramaje_Lavado").Value
    End If
    
    If IsNull(Rs("Observaciones").Value) Then
        Me.txtObservaciones.Text = ""
    Else
        Me.txtObservaciones.Text = Rs("Observaciones").Value
    End If
    
    If IsNull(Rs("Flg_Status_Apro").Value) Then
        Me.cboFlg_Status_Apro.ListIndex = -1
    Else
        Call BuscaCombo(Rs("Flg_Status_Apro").Value, 2, cboFlg_Status_Apro)
    End If
   
    txtKilosRechazados.Text = Rs("Kgs_Rechazados").Value
End If

Set RsTelas = New ADODB.Recordset
RsTelas.Open "select gramaje_acab,ancho_acab,encog_Ancho,encog_largo from tx_tela where cod_tela='" & vCod_Tela & "'", cConnect, adOpenStatic
If RsTelas.RecordCount Then
    TxtGramPrg.Text = RsTelas("gramaje_Acab")
    TxtAnchoPrg.Text = RsTelas("ancho_acab")
    TxtEncAncPrg.Text = RsTelas("encog_Ancho")
    TxtEncLarPrg.Text = RsTelas("encog_largo")
End If

Set Rs = Nothing
Set RsTelas = Nothing
Exit Sub

hand:
Set Rs = Nothing
Set RsTelas = Nothing

ErrorHandler err, "Form_load"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "ACEPTAR"
        SALVAR_DATOS
        
    Case "CANCELAR"
        Unload Me
End Select
End Sub

Sub SALVAR_DATOS()
Dim Reg As ADODB.Recordset
On Error GoTo hand

If Not IsNumeric(TxtEncogA) Then
    MsgBox "Cantidad de encogimiento de ancho erronea", vbExclamation, Me.Caption
    TxtEncogA.SetFocus
    Exit Sub
End If

If Not IsNumeric(TxtEncogL) Then
    MsgBox "Cantidad de encogimiento de largp erronea", vbExclamation, Me.Caption
    TxtEncogL.SetFocus
    Exit Sub
End If

If CboTipo.ListIndex = -1 Then
    MsgBox "Ingrese el tipo de encogimiento", vbExclamation, Me.Caption
    CboTipo.SetFocus
    Exit Sub
End If

If Trim(TxtElongacionLargo) = "" Then
    TxtElongacionLargo = "0"
End If

If Trim(TxtElongacionAncho) = "" Then
    TxtElongacionAncho = "0"
End If

If Right(cboFlg_Status_Apro.Text, 1) = "D" And txtKilosRechazados.Text <= 0 Then
    Aviso "Cantidad de Kilos Rechazados es Obligatorio", 1
    txtKilosRechazados.SetFocus
    Exit Sub
End If

Set Reg = New ADODB.Recordset
Reg.ActiveConnection = cConnect
Reg.CursorLocation = adUseClient
Reg.CursorType = adOpenStatic

Reg.Open "UP_MAN_ORDTRA_TELAS '" & vCod_TipOrdTra & "','" & _
                                vCod_OrdTra & "','" & _
                                vCod_Tela & "','" & _
                                vCod_Comb & "','" & _
                                vCod_Color & "'," & _
                                TxtGramaje.Text & "," & _
                                TxtAncho.Text & "," & _
                                TxtEncogA.Text & "," & _
                                TxtEncogL.Text & ",'" & _
                                vusu & "','" & _
                                Trim(Right(Me.CboTipo, 1)) & "'," & _
                                TxtRevirado.Text & "," & _
                                IIf(Trim(Me.txtGramaje_Lavado.Text) = "", 0, Me.txtGramaje_Lavado.Text) & "," & _
                                IIf(Trim(TxtGramajePP.Text) = "", 0, TxtGramajePP.Text) & "," & _
                                IIf(Trim(TxtAnchoPP.Text) = "", 0, TxtAnchoPP.Text) & "," & _
                                IIf(Trim(TxtEncogAPP.Text) = "", 0, TxtEncogAPP.Text) & "," & _
                                TxtEncogLPP.Text & "," & _
                                TxtReviradoPP.Text & "," & _
                                IIf(Trim(Me.txtGramaje_LavadoPP.Text) = "", 0, Me.txtGramaje_LavadoPP.Text) & ",'" & _
                                Me.txtObservaciones.Text & "','" & _
                                Right(Me.cboFlg_Status_Apro.Text, 1) & "','" & _
                                CDbl(TxtElongacionLargo) & "','" & TxtCod_MotRechazo & "','" & _
                                CDbl(TxtElongacionAncho) & "'," & CDbl(txtKilosRechazados.Text) & ""
Set Reg = Nothing
Unload Me

Exit Sub
hand:
ErrorHandler err, "Salvar_Datos"
End Sub



Private Sub TxtAncho_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AVANZA 13
    End If
End Sub

Private Sub TxtAncho_LostFocus()
    If Not IsNumeric(Trim(TxtAncho.Text)) Then
        MsgBox "El valor ingresado no es correcto", vbCritical, Me.Caption
        TxtAncho.Text = ""
        TxtAncho.SetFocus
    End If
End Sub

Private Sub TxtAnchoPP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AVANZA 13
    End If
End Sub

Private Sub TxtAnchoPP_LostFocus()
    If Not IsNumeric(Trim(TxtAnchoPP.Text)) Then
        MsgBox "El valor ingresado no es correcto", vbCritical, Me.Caption
        TxtAnchoPP.Text = ""
        TxtAnchoPP.SetFocus
    End If
End Sub

Private Sub TxtCod_MotRechazo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call AYUDA_MOTRECHAZO
End If
End Sub

Private Sub TxtEncogA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AVANZA 13
    End If
End Sub

Private Sub TxtEncogA_LostFocus()
    If Not IsNumeric(Trim(TxtEncogA.Text)) Then
        MsgBox "El valor ingresado no es correcto", vbCritical, Me.Caption
        TxtEncogA.Text = ""
        TxtEncogA.SetFocus
    End If
End Sub

Private Sub TxtEncogAPP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AVANZA 13
    End If

End Sub

Private Sub TxtEncogAPP_LostFocus()
    If Not IsNumeric(Trim(TxtEncogAPP.Text)) Then
        MsgBox "El valor ingresado no es correcto", vbCritical, Me.Caption
        TxtEncogAPP.Text = ""
        TxtEncogAPP.SetFocus
    End If
End Sub

Private Sub TxtEncogL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AVANZA 13
    End If

End Sub

Private Sub TxtEncogL_LostFocus()
    If Not IsNumeric(Trim(TxtEncogL.Text)) Then
        MsgBox "El valor ingresado no es correcto", vbCritical, Me.Caption
        TxtEncogL.Text = ""
        TxtEncogL.SetFocus
    End If
End Sub

Private Sub TxtEncogLPP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AVANZA 13
    End If

End Sub

Private Sub TxtEncogLPP_LostFocus()
    If Not IsNumeric(Trim(TxtEncogLPP.Text)) Then
        MsgBox "El valor ingresado no es correcto", vbCritical, Me.Caption
        TxtEncogLPP.Text = ""
        TxtEncogLPP.SetFocus
    End If
End Sub

'Private Sub TxtEncogA_KeyPress(KeyAscii As Integer)
'    Call SoloNumeros(TxtEncogA, KeyAscii, True, 2, 6)
'End Sub

'Private Sub TxtEncogL_KeyPress(KeyAscii As Integer)
'    Call SoloNumeros(TxtEncogL, KeyAscii, True, 2, 6)
'End Sub

Private Sub TxtGramaje_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AVANZA 13
    Else
        Call SoloNumeros(TxtGramaje, KeyAscii, False, 0, 5)
    End If
End Sub

Private Sub txtGramaje_Lavado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AVANZA 13
    Else
        Call SoloNumeros(Me.txtGramaje_Lavado, KeyAscii, False, 0, 5)
    End If
End Sub

Private Sub txtGramaje_Lavado_LostFocus()
    If Trim(Me.txtGramaje_Lavado.Text) = "" Then
        Me.txtGramaje_Lavado = "0"
    End If
End Sub

Private Sub txtGramaje_LavadoPP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AVANZA 13
    Else
        Call SoloNumeros(Me.txtGramaje_Lavado, KeyAscii, False, 0, 5)
    End If
End Sub

Private Sub TxtGramajePP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AVANZA 13
    Else
        Call SoloNumeros(TxtGramajePP, KeyAscii, False, 0, 5)
    End If
End Sub

Private Sub TxtRevirado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AVANZA 13
    End If
End Sub

Private Sub TxtRevirado_LostFocus()
    If Not IsNumeric(Trim(TxtRevirado.Text)) Then
        MsgBox "El valor ingresado no es correcto", vbCritical, Me.Caption
        TxtRevirado.Text = ""
        TxtRevirado.SetFocus
    End If
End Sub

Private Sub TxtReviradoPP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AVANZA 13
    End If
End Sub

Private Sub TxtReviradoPP_LostFocus()
    If Not IsNumeric(Trim(TxtReviradoPP.Text)) Then
        MsgBox "El valor ingresado no es correcto", vbCritical, Me.Caption
        TxtReviradoPP.Text = ""
        TxtReviradoPP.SetFocus
    End If
End Sub

Sub AYUDA_MOTRECHAZO()
    Dim oTipo As New frmBusqGeneral2
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    Codigo = ""
    Descripcion = ""
    oTipo.sQuery = "SELECT Cod_MotRechazo AS 'Código', Des_MotRechazo as 'Descripción' FROM TX_MOTIVO_RECHAZO"
    oTipo.CARGAR_DATOS
    oTipo.Show 1
    If Codigo <> "" Then
        Me.TxtCod_MotRechazo.Text = Trim(Codigo)
        Me.TxtDes_MotRechazo.Text = Trim(Descripcion)
        Codigo = "": Descripcion = ""
        'txtCod_TemCli.SetFocus
    End If
    Set oTipo = Nothing
    Set Rs = Nothing

End Sub
