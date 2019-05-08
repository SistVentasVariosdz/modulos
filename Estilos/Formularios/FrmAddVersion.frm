VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmAddVersion 
   Caption         =   "Versiones"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkWipProto 
      Caption         =   "Seguimiento WipProto"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2520
      TabIndex        =   22
      Top             =   4320
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmAddVersion.frx":0000
      Orientacion     =   0
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
      Height          =   4230
      Left            =   0
      TabIndex        =   23
      Tag             =   "Detail"
      Top             =   0
      Width           =   7515
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   380
         Left            =   120
         TabIndex        =   36
         Top             =   2550
         Width           =   7305
         Begin VB.ComboBox dbcSolicitud 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   0
            Width           =   2025
         End
         Begin VB.ComboBox dbcPrenda 
            Height          =   315
            Left            =   5040
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   0
            Width           =   2115
         End
         Begin VB.Label Label5 
            Caption         =   "Solicitud"
            Height          =   195
            Left            =   60
            TabIndex        =   38
            Top             =   30
            Width           =   675
         End
         Begin VB.Label Label6 
            Caption         =   "Nivel Prenda"
            Height          =   255
            Left            =   3960
            TabIndex        =   37
            Top             =   60
            Width           =   945
         End
      End
      Begin VB.TextBox TxtCodigo 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1410
         MaxLength       =   2
         TabIndex        =   0
         Top             =   240
         Width           =   690
      End
      Begin VB.CommandButton cmdIcono1 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   6930
         TabIndex        =   4
         Top             =   645
         Width           =   420
      End
      Begin VB.TextBox TxtIcono1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4020
         TabIndex        =   3
         Top             =   660
         Width           =   2910
      End
      Begin VB.TextBox txtDescripcion 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   5205
      End
      Begin VB.Frame Frame2 
         Height          =   1290
         Left            =   120
         TabIndex        =   30
         Top             =   2880
         Width           =   7275
         Begin VB.TextBox Text2 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   15
            Top             =   255
            Width           =   1965
         End
         Begin VB.TextBox TxtEstilo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   900
            TabIndex        =   14
            Top             =   255
            Width           =   690
         End
         Begin VB.TextBox TxtVersion 
            BackColor       =   &H00FFFFFF&
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
            Left            =   4410
            MaxLength       =   5
            TabIndex        =   16
            Top             =   255
            Width           =   645
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Copiado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   0
            TabIndex        =   13
            Top             =   30
            Width           =   1365
         End
         Begin VB.TextBox TxtDesVersion 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5070
            TabIndex        =   17
            Top             =   255
            Width           =   2115
         End
         Begin VB.OptionButton OptEstilo 
            Caption         =   "Estilo"
            Enabled         =   0   'False
            Height          =   210
            Left            =   75
            TabIndex        =   33
            Top             =   300
            Value           =   -1  'True
            Width           =   780
         End
         Begin VB.OptionButton OptPlantilla 
            Caption         =   "Cliente"
            Enabled         =   0   'False
            Height          =   195
            Left            =   90
            TabIndex        =   32
            Top             =   660
            Width           =   780
         End
         Begin VB.TextBox TxtAbr_Cliente 
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   315
            Left            =   900
            TabIndex        =   18
            Top             =   600
            Width           =   720
         End
         Begin VB.TextBox TxtNom_Cliente 
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1635
            TabIndex        =   19
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtNum_Plantilla 
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   315
            Left            =   4425
            TabIndex        =   20
            Top             =   600
            Width           =   630
         End
         Begin VB.TextBox txtNom_Plantilla 
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5085
            TabIndex        =   21
            Top             =   600
            Width           =   2100
         End
         Begin VB.CheckBox ChkCopiaMatriz 
            Caption         =   "Copiar Matriz Colores"
            Height          =   255
            Left            =   5400
            TabIndex        =   31
            Top             =   960
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Version"
            Height          =   315
            Index           =   4
            Left            =   3720
            TabIndex        =   35
            Top             =   285
            Width           =   555
         End
         Begin VB.Label Label8 
            Caption         =   "Plantilla"
            Height          =   225
            Left            =   3720
            TabIndex        =   34
            Top             =   645
            Width           =   585
         End
      End
      Begin VB.ComboBox dbcTipo 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   660
         Width           =   1500
      End
      Begin VB.Frame Frame3 
         Caption         =   "Mano de Obra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   120
         TabIndex        =   24
         Top             =   1470
         Width           =   7275
         Begin VB.TextBox TxtEfi 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1320
            MaxLength       =   20
            TabIndex        =   9
            Text            =   "0.00"
            Top             =   600
            Width           =   1005
         End
         Begin VB.TextBox TxtMin 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3780
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   240
            Width           =   1020
         End
         Begin VB.TextBox TxtCorte 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1320
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   240
            Width           =   1005
         End
         Begin VB.TextBox TxtAcabado 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6120
            TabIndex        =   8
            Text            =   "0.00"
            Top             =   240
            Width           =   1035
         End
         Begin VB.TextBox TxtRectilineos 
            Height          =   315
            Left            =   6120
            MaxLength       =   5
            TabIndex        =   10
            Top             =   600
            Width           =   1035
         End
         Begin VB.Label Label2 
            Caption         =   "Efic.Costura (%)"
            Height          =   225
            Index           =   1
            Left            =   130
            TabIndex        =   29
            Top             =   680
            Width           =   1245
         End
         Begin VB.Label Label2 
            Caption         =   "Min.Costura"
            Height          =   315
            Index           =   0
            Left            =   2670
            TabIndex        =   28
            Top             =   300
            Width           =   885
         End
         Begin VB.Label Label3 
            Caption         =   "Min. Corte"
            Height          =   225
            Left            =   130
            TabIndex        =   27
            Top             =   320
            Width           =   825
         End
         Begin VB.Label Label4 
            Caption         =   "Min.Acabado"
            Height          =   255
            Left            =   5010
            TabIndex        =   26
            Top             =   270
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Nro.Espec.Rectilineos:"
            Height          =   255
            Left            =   4320
            TabIndex        =   25
            Top             =   680
            Width           =   1695
         End
      End
      Begin VB.TextBox TxtDes_TelaCliente 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1365
         TabIndex        =   5
         Top             =   1100
         Width           =   5985
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
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
         Left            =   135
         TabIndex        =   42
         Tag             =   "Description:"
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label2 
         Caption         =   "Imagen:"
         Height          =   255
         Index           =   3
         Left            =   3180
         TabIndex        =   41
         Top             =   700
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo:"
         Height          =   285
         Left            =   120
         TabIndex        =   40
         Top             =   700
         Width           =   585
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Des. Tela Cliente:"
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
         Left            =   105
         TabIndex        =   39
         Tag             =   "Description:"
         Top             =   1150
         Width           =   1245
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   6960
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmAddVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Reg As New ADODB.Recordset
Public vCodEstPro As String
Public Codigo, Descripcion As String
Public Cliente As String
Public sCod_Cliente As String
Public sNom_Cliente As String

Dim Estado As String

Public TipoBusq As Integer
Public vCodCliente As String
Public vCodTemporada As String
Public vNumCotizacion As Integer

Public sCod_Estcli As String
'Variables creadas por AHSP
Dim sTipo As String
Dim StrSQL As String
Dim vCod As String
'Dim vDes_Tela As String

Public Tipo_Busqueda As String
Public vDes_Tela As String


Public Function VALIDA_DATOS() As Boolean

    VALIDA_DATOS = True
    If sTipo <> "B" Then
        If Trim(TxtCodigo.Text) = "" Then
            Call MsgBox("El código no puede estar vacio. Sirvase verificar", vbInformation, "Versiones")
            TxtCodigo.Text = ""
            TxtCodigo.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
    
        If Estado = "NUEVO" Then
            StrSQL = "Select count(*) from es_estprover where COD_ESTPRO='" & vCodEstPro & "' AND Cod_Version='" & Trim(TxtCodigo.Text) & "'"
            If DevuelveCampo(StrSQL, cCONNECT) > 0 Then
                Call MsgBox("El código de versión ya se encuentra registrado. Sirvase verificar", vbInformation, "Versiones")
                TxtCodigo.SetFocus
                VALIDA_DATOS = False
                Exit Function
            End If
        End If
        
        If Trim(txtDescripcion.Text) = "" Then
            Call MsgBox("La descripción no puede estar vacia. Sirvase verificar", vbInformation, "Versiones")
            txtDescripcion.Text = ""
            txtDescripcion.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If

    Else
    
        StrSQL = "Select count(*) from es_estprocol where where COD_ESTPRO='" & vCodEstPro & "' AND Cod_Version='" & Trim(TxtCodigo.Text) & "'"
        If DevuelveCampo(StrSQL, cCONNECT) > 0 Then
            Call MsgBox("Esta versión esta relacionada con colores propios. Sirvase verificar", vbInformation, "Versiones")
            VALIDA_DATOS = False
            Exit Function
        End If

        StrSQL = "Select count(*) from es_estcompcol where COD_ESTPRO='" & vCodEstPro & "' AND Cod_Version='" & Trim(TxtCodigo.Text) & "'"
        If DevuelveCampo(StrSQL, cCONNECT) > 0 Then
            Call MsgBox("Esta versión esta relacionada con colores. Sirvase verificar", vbInformation, "Versiones")
            VALIDA_DATOS = False
            Exit Function
        End If
        
        StrSQL = "Select count(*) from es_estprocomp where COD_ESTPRO='" & vCodEstPro & "' AND Cod_Version='" & Trim(TxtCodigo.Text) & "'"
        If DevuelveCampo(StrSQL, cCONNECT) > 0 Then
            Call MsgBox("Esta versión esta relacionada con componentes propios. Sirvase verificar", vbInformation, "Versiones")
            VALIDA_DATOS = False
            Exit Function
        End If
    
        StrSQL = "Select count(*) from es_aplan_items  where COD_ESTPRO='" & vCodEstPro & "' AND Cod_Version='" & Trim(TxtCodigo.Text) & "'"
        If DevuelveCampo(StrSQL, cCONNECT) > 0 Then
            Call MsgBox("Esta versión tiene ediciones Avios. Sirvase verificar", vbInformation, "Versiones")
            VALIDA_DATOS = False
            Exit Function
        End If
        
        StrSQL = "Select count(*) from es_aplan_telas   where COD_ESTPRO='" & vCodEstPro & "' AND Cod_Version='" & Trim(TxtCodigo.Text) & "'"
        If DevuelveCampo(StrSQL, cCONNECT) > 0 Then
            Call MsgBox("Esta versión tiene ediciones Telas. Sirvase verificar", vbInformation, "Versiones")
            VALIDA_DATOS = False
            Exit Function
        End If
        
    End If
    
End Function

Sub Accion(pTipo As String, pVersion As String, pDescripcion As String, pIcono As String, pMinimo As Double, pEfi As Double, Corte As Double, Acabado As Double, Solicitud As String, nivel As String, Optional EsAccion As Boolean = False)
On Error GoTo hand
    Set Reg = Nothing
    Reg.CursorLocation = adUseClient
    Reg.Open "UP_Es_EstProVer '" & pTipo & "','" & vCodEstPro & "','" & pVersion & "','" & pDescripcion & "','" & pIcono & "'," & pMinimo & "," & pEfi & ",'" & Mid(dbcTipo.Text, 1, 1) & "'," & Corte & "," & Acabado & ",'" & Solicitud & "','" & nivel & "','" & Trim(TxtRectilineos.Text) & "','" & Me.TxtDes_TelaCliente & "','" & Me.sCod_Cliente & "','" & Me.vCodTemporada & "','N','" & vusu & "'", cCONNECT
    If pTipo = "I" Or pTipo = "A" Then
        If pTipo = "I" Then
            Call NUEVO_DATO_DESTELA("2")
        Else
            Call VERIFICA_ULTIMO
        End If
    End If
Exit Sub
hand:

ErrorHandler Err, "Accion"
'Accion "V", "", "", "", 0, 0, 0, 0, "", "", False
End Sub

Sub Copiado()
On Error GoTo hand

If OptEstilo.Value = True Then
    StrSQL = "up_copia_version_estilo '" & TxtEstilo & "','" & TxtVersion & "','" & vCodEstPro & "','" & TxtCodigo & "','" & IIf(ChkCopiaMatriz, "S", "N") & "'"
Else
    StrSQL = "up_copia_estilo_version_componentes_desde_plantilla '" & vCodEstPro & "','" & TxtCodigo.Text & "'," & txtNum_Plantilla.Text & ""
End If

ExecuteCommandSQL cCONNECT, StrSQL
Exit Sub
hand:
    MsgBox Err.Description, vbCritical, "Copiado"
End Sub

Function ValidaEstilo() As Boolean
Dim temp As String
temp = Trim(DevuelveCampo("Select dbo.uf_devuelvecodigo(5," & TxtEstilo & ")", cCONNECT))
   TxtEstilo = temp
   If ExisteCampo("Cod_EstPro", "es_estpro", temp, cCONNECT, True) Then
        ValidaEstilo = True
        Text2 = DevuelveCampo("select Des_estpro from es_estpro where cod_estpro='" & temp & "'", cCONNECT)
        TxtVersion.SetFocus
    Else
        ValidaEstilo = False
        Text2 = ""
  End If
End Function

Private Sub Check1_Click()
If Check1.Value = 0 Then
    TxtEstilo.Enabled = False
    TxtVersion.Enabled = False
    OptEstilo.Enabled = False
    OptPlantilla.Enabled = False
    ChkCopiaMatriz.Enabled = False
    TxtEstilo = ""
    TxtEstilo = ""
    Text2 = ""
Else
    TxtEstilo.Enabled = True
    TxtVersion.Enabled = True
    OptEstilo.Enabled = True
    OptPlantilla.Enabled = True
    ChkCopiaMatriz.Enabled = True
    TxtEstilo.SetFocus
End If
End Sub

Private Sub cmdFirst_Click()
    If Not Reg.BOF Then Reg.MoveFirst
End Sub

Private Sub cmdIcono1_Click()
With cd
    .ShowSave
    Me.TxtIcono1 = .FileName
    SendKeys "{TAB}"
End With

End Sub

Private Sub cmdLast_Click()
    If Not Reg.EOF Then Reg.MoveLast
End Sub

Private Sub cmdNext_Click()
If Not Reg.EOF Then Reg.MoveNext
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo hand
If Not Reg.BOF Then Reg.MovePrevious
Exit Sub
hand:
ErrorHandler Err, "cmdPrevious_Click"
End Sub

Private Sub dbcPrenda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub dbcSolicitud_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub dbcTipo_Click()
    If dbcTipo.ListIndex = 2 Then
        Frame4.Visible = True
        Call DefaultCombos("select cod_tipsolcos from es_tipsolcos where flg_default='*'", dbcSolicitud)
        Call DefaultCombos("select cod_nivpda from es_nivpda where flg_default='*'", dbcPrenda)
    Else
        Frame4.Visible = False
    End If
End Sub

Private Sub Form_Load()
Me.Caption = Me.Caption & " del estilo: " & vCodEstPro
Call FormSet(Me)

'dbcTipo.AddItem "Producción", 0
'dbcTipo.AddItem "Costeo", 1
CARGA_COMBOS
Check1_Click
'Habilita False
End Sub

'Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
'On Error GoTo hand
'Select Case ActionName
'    Case "ADICIONAR"
'        Limpia
'        Habilita True
'        Estado = "NUEVO"
'        Call NUEVO_DATO_DESTELA("1")
'        Me.TxtCodigo.SetFocus
'        Check1.Enabled = True
'        sTipo = "I"
'End Select
'Exit Sub
'hand:
'ErrorHandler Err, "MantFunc1_ActionClick"
'End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call Grabar_Version
Case "SALIR"
    Unload Me
End Select
End Sub

Private Sub optEstilo_Click()
    TxtAbr_Cliente.Enabled = False
    TxtNom_Cliente.Enabled = False
    txtNum_Plantilla.Enabled = False
    txtNom_Plantilla.Enabled = False
    TxtAbr_Cliente.BackColor = &H80000000
    TxtNom_Cliente.BackColor = &H80000000
    txtNum_Plantilla.BackColor = &H80000000
    txtNom_Plantilla.BackColor = &H80000000
    
    TxtEstilo.Enabled = True
    Text2.Enabled = True
    TxtVersion.Enabled = True
    TxtDesVersion.Enabled = True
    TxtEstilo.BackColor = &H80000005
    Text2.BackColor = &H80000005
    TxtVersion.BackColor = &H80000005
    TxtDesVersion.BackColor = &H80000005
    
    TxtEstilo.SetFocus
End Sub

Private Sub OptPlantilla_Click()
    TxtEstilo.Enabled = False
    Text2.Enabled = False
    TxtVersion.Enabled = False
    TxtDesVersion.Enabled = False
    TxtEstilo.BackColor = &H80000000
    Text2.BackColor = &H80000000
    TxtVersion.BackColor = &H80000000
    TxtDesVersion.BackColor = &H80000000

    TxtAbr_Cliente.Enabled = True
    TxtNom_Cliente.Enabled = True
    txtNum_Plantilla.Enabled = True
    txtNom_Plantilla.Enabled = True
    TxtAbr_Cliente.BackColor = &H80000005
    TxtNom_Cliente.BackColor = &H80000005
    txtNum_Plantilla.BackColor = &H80000005
    txtNom_Plantilla.BackColor = &H80000005
    
    If Trim(TxtAbr_Cliente.Text) = "" Then
        TxtAbr_Cliente.SetFocus
    Else
        txtNum_Plantilla.SetFocus
    End If
End Sub

Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(TxtAbr_Cliente.Text) = "" Then
            BUSCA_CLIENTE (3)
        Else
            BUSCA_CLIENTE (1)
        End If
    End If
End Sub

Private Sub TxtAcabado_GotFocus()
SelectionText TxtAcabado
End Sub

Private Sub TxtAcabado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtAcabado, KeyAscii, True, 2)
End If
End Sub

Private Sub TxtCodigo_GotFocus()
SelectionText TxtCodigo
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(TxtCodigo.Text) <> "" Then txtDescripcion.SetFocus
    End If
End Sub

Private Sub TxtCorte_GotFocus()
SelectionText TxtCorte
End Sub

Private Sub TxtCorte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtCorte, KeyAscii, True, 2)
End If
End Sub

Private Sub TxtDes_TelaCliente_GotFocus()
vDes_Tela = Trim(TxtDes_TelaCliente.Text)
End Sub

Private Sub TxtDes_TelaCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtDes_TelaCliente_LostFocus()
If Me.Tipo_Busqueda <> "3" Then
    If Me.TxtDes_TelaCliente <> vDes_Tela Then
        MsgBox "Tipo de busqueda no permite modificar en Est. Cliente Temporada"
        TxtDes_TelaCliente.Text = vDes_Tela
    End If
End If
End Sub

Private Sub txtDescripcion_GotFocus()
SelectionText txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtEfi_GotFocus()
SelectionText TxtEfi
End Sub

Private Sub TxtIcono1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtMin_GotFocus()
SelectionText TxtMin
End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(TxtNom_Cliente.Text) = "" Then
            BUSCA_CLIENTE (3)
        Else
            BUSCA_CLIENTE (2)
        End If
    End If
End Sub

Private Sub TxtEfi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    SoloNumeros TxtEfi, KeyAscii, True, 2, 3
End If
End Sub


Private Sub TxtEstilo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxtEstilo.Text) = "" Then
        Exit Sub
    End If

    If ValidaEstilo Then
    End If
Else
    SoloNumeros TxtEstilo, KeyAscii, False, 0, 5
End If

End Sub


Private Sub TxtMin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    SoloNumeros TxtMin, KeyAscii, True, 3, 3
End If
End Sub


Private Sub CARGA_COMBOS()
LlenaCombo dbcSolicitud, "select des_tipsolcos+space(100)+cod_tipsolcos from es_tipsolcos order by des_tipsolcos", cCONNECT
LlenaCombo dbcPrenda, "select des_nivpda+space(100)+cod_nivpda from es_nivpda order by des_nivpda", cCONNECT
LlenaCombo dbcTipo, "select Tip_Version + space(1) + descripcion from Es_TiposVersion order by descripcion desc ", cCONNECT
End Sub

Public Sub DefaultCombos(sSQl As String, combo As ComboBox)
Dim Rsd As New ADODB.Recordset

Rsd.Open sSQl, cCONNECT, adOpenStatic
If Rsd.RecordCount > 0 Then
    Call BuscaCombo(Rsd(0), 2, combo)
Else
    combo.ListIndex = -1
End If
Set Rsd = Nothing
End Sub

Function ValidaVersion() As Boolean
Dim temp As String
temp = Trim(DevuelveCampo("select des_Version from es_Estprover where cod_estpro='" & Trim(TxtEstilo.Text) & "' and cod_version='" & Trim(TxtVersion.Text) & "'", cCONNECT))
If temp <> "" Then
    TxtDesVersion = temp
    ValidaVersion = True
    FunctButt1.SetFocus
Else
    TxtVersion = ""
    TxtDesVersion = ""
    ValidaVersion = False
End If
End Function


Private Sub txtNom_Plantilla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtNom_Plantilla.Text) = "" Then
            BUSCA_PLANTILLA (3)
        Else
            BUSCA_PLANTILLA (2)
        End If
    End If
End Sub

Private Sub txtNum_Plantilla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtNum_Plantilla.Text) = "" Then
            BUSCA_PLANTILLA (3)
        Else
            BUSCA_PLANTILLA (1)
        End If
    End If
End Sub

Private Sub TxtRectilineos_GotFocus()
SelectionText TxtRectilineos
End Sub

Private Sub TxtRectilineos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtRectilineos, KeyAscii, True, 2)
End If
End Sub

Private Sub TxtVersion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxtEstilo.Text) = "" Then
        Exit Sub
    End If
    If ValidaVersion = False Then
        frmBusqGeneral.sQuery = "select cod_version as Codigo,des_Version as Descripcion from es_Estprover where cod_estpro='" & TxtEstilo.Text & "'"
        Set frmBusqGeneral.oParent = Me
        frmBusqGeneral.Cargar_Datos
        frmBusqGeneral.Show 1
        TxtVersion = Codigo
        TxtDesVersion = Descripcion
        FunctButt1.SetFocus
    End If
End If
End Sub

Public Sub BUSCA_CLIENTE(tipo As Integer)
    Select Case tipo
        Case 1:
                    StrSQL = "SELECT nom_cliente FROM Tg_cliente WHERE abr_cliente = '" & Trim(Me.TxtAbr_Cliente.Text) & "'"
                    Me.TxtNom_Cliente.Text = Trim(DevuelveCampo(StrSQL, cCONNECT))
                    Me.txtNum_Plantilla.SetFocus
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim Rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        oTipo.sQuery = "SELECT abr_cliente AS 'Código', nom_cliente AS 'Descripción' FROM tg_cliente where nom_cliente like '%" & Trim(TxtNom_Cliente.Text) & "%' order by abr_cliente"
                    Else
                        oTipo.sQuery = "SELECT abr_cliente AS 'Código', nom_cliente AS 'Descripción' FROM tg_cliente order by abr_cliente"
                    End If
                    
                    oTipo.Cargar_Datos
                    oTipo.Show 1
                    If Codigo <> "" Then
                         Me.TxtAbr_Cliente.Text = Trim(Codigo)
                         Me.TxtNom_Cliente.Text = Trim(Descripcion)
                         Codigo = "": Descripcion = ""
                         Me.txtNum_Plantilla.SetFocus
                    End If
                    Set oTipo = Nothing
                    Set Rs = Nothing
    End Select
    
End Sub

Public Sub BUSCA_PLANTILLA(tipo As Integer)
Dim sCliente As String

StrSQL = "select cod_cliente from tg_cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
vCodCliente = DevuelveCampo(StrSQL, cCONNECT)

    Select Case tipo
        Case 1:
                    StrSQL = "select nombre from  Es_Plantillas_Componentes WHERE Cod_Cliente='" & vCodCliente & "' and num_plantilla = " & Me.txtNum_Plantilla.Text
                    Me.txtNom_Plantilla.Text = Trim(DevuelveCampo(StrSQL, cCONNECT))
                    Me.FunctButt1.SetFocus
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim Rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        oTipo.sQuery = "SELECT Num_Plantilla AS 'Código', nombre AS 'Descripción' FROM Es_Plantillas_Componentes where Cod_Cliente='" & vCodCliente & "' and nombre like '%" & Trim(txtNom_Plantilla.Text) & "%' order by abr_cliente"
                    Else
                        oTipo.sQuery = "SELECT Num_Plantilla AS 'Código', nombre AS 'Descripción' FROM Es_Plantillas_Componentes Where Cod_Cliente='" & vCodCliente & "' order by nombre"
                    End If
                    
                    oTipo.Cargar_Datos
                    oTipo.Show 1
                    If Codigo <> "" Then
                         Me.txtNum_Plantilla.Text = Trim(Codigo)
                         Me.txtNom_Plantilla.Text = Trim(Descripcion)
                         Codigo = "": Descripcion = ""
                         Me.FunctButt1.SetFocus
                    End If
                    Set oTipo = Nothing
                    Set Rs = Nothing
    End Select
End Sub

Sub NUEVO_DATO_DESTELA(ByVal opcion As String)
Dim Rs As ADODB.Recordset
On Error GoTo hand

Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseClient

'If Tipo_Busqueda <> "3" Then
'    MsgBox "Des. Tela Cliente no se puede editar"
'Exit Sub
'End If

If Tipo_Busqueda = "3" Then
    'If IIf(dbcTipo.ListIndex = 1, "C", "P") = "P" Then
    If Mid(dbcTipo.Text, 1, 1) = "P" Then
        StrSQL = "sm_devuelvetela_ESTCLITEM '" & opcion & "','" & Me.sCod_Cliente & "','" & vCodTemporada & "','" & sCod_Estcli & "','" & Me.TxtDes_TelaCliente & "'"
        If opcion = "1" Then
            TxtDes_TelaCliente = DevuelveCampo(StrSQL, cCONNECT)
        'Else
            'rs.Open strSQL, cCONNECT
            'Set rs = Nothing
        End If
    End If
End If
Exit Sub
hand:
    ErrorHandler Err, "SALVAR_DATOS"
    Set Rs = Nothing
End Sub

Sub VERIFICA_ULTIMO()
StrSQL = "MUESTRA_MAX_VERSION '" & vCodEstPro & "'"
If Trim(TxtCodigo) = Trim(DevuelveCampo(StrSQL, cCONNECT)) Then
    Call NUEVO_DATO_DESTELA("2")
End If
End Sub

Sub Grabar_Version()
On Error GoTo errGrabar
If Len(Trim(TxtEfi)) = 0 Then TxtEfi = "0"
If Len(Trim(TxtMin)) = 0 Then TxtMin = "0"
If Me.VALIDA_DATOS = True Then
    If Check1.Value = 1 Then
        If OptEstilo.Value Then
            If Not ValidaEstilo Then Exit Sub
            If Trim(TxtVersion) = "" Then
                MsgBox "Debe llenar la version", vbCritical
                Exit Sub
            End If
        Else
            If Trim(TxtAbr_Cliente.Text) = "" Then
                MsgBox "Seleccione el Cliente", vbInformation, "Aviso"
                TxtAbr_Cliente.SetFocus
                Exit Sub
            End If
            
            If Trim(txtNum_Plantilla.Text) = "" Then
                MsgBox "Seleccione el Numero de Plantilla a Copiar", vbInformation, "Aviso"
                txtNum_Plantilla.SetFocus
                Exit Sub
            End If
            
            If Not ExisteCampo("Num_Plantilla", "Es_Plantillas_Componentes", txtNum_Plantilla.Text, cCONNECT, True) Then
                MsgBox "Plantilla no existe", vbInformation, "Aviso"
                txtNum_Plantilla.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    Accion "I", Me.TxtCodigo, Me.txtDescripcion, Me.TxtIcono1, Me.TxtMin, Me.TxtEfi, Me.TxtCorte, Me.TxtAcabado, Right(dbcSolicitud, 3), Right(dbcPrenda, 2), True
    
    'StrSQL = "Es_Actualiza_Version_Costeo_Estilo '" & Me.sCod_Cliente & "','" & Me.vCodTemporada & "','" & Me.sCod_Estcli & "','" & Me.vCodEstPro & "','" & Trim(TxtCodigo.Text) & "','I',0,'" & vusu & "'"
    'Call ExecuteSQL(cCONNECT, StrSQL)
    
    'AGREGA AL SEGUIMIENTO DEL WIPPROTO
    'if ChkWipProto then
    '    Call ExecuteCommandSQL(cCONNECT, "es_incorpora_estilo_version_al_wip '" & vCodEstPro & "','" & TxtCodigo.Text & "','" & vusu & "','" & format(date,"dd/mm/yyyy)" & "','" & sCod_Cliente & "','" & vCodTemporada & "','" & sCod_Estcli & "'")
    'endif
    
    If Check1.Value = 1 Then
        Copiado
    Else
        Load FrmAddComponentes
        FrmAddComponentes.vCod_estPro = Me.vCodEstPro
        FrmAddComponentes.Txtcod_EstPro = Me.vCodEstPro
        FrmAddComponentes.TxtDes_EstPro = DevuelveCampo("select des_estpro from es_estpro where cod_estpro ='" & Me.vCodEstPro & "'", cCONNECT)
        FrmAddComponentes.vCod_Version = TxtCodigo.Text
        FrmAddComponentes.TxtCod_Version = TxtCodigo.Text
        FrmAddComponentes.txtDes_Version = Trim(txtDescripcion)
        FrmAddComponentes.CARGA_GRID
        FrmAddComponentes.Show vbModal
        Set FrmAddComponentes = Nothing
    End If
    
    Unload Me
End If
Exit Sub
errGrabar:
    MsgBox Err.Description, vbCritical, "Grabar Version"
End Sub
