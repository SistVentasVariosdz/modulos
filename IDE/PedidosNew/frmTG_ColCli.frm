VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form frmTG_ColCli 
   Caption         =   "Creación de Color"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Add Color"
   Begin VB.Frame fraAddPresent 
      Caption         =   "Adicionar Presentaciones"
      Height          =   2145
      Left            =   90
      TabIndex        =   18
      Top             =   1260
      Visible         =   0   'False
      Width           =   5730
      Begin VB.TextBox txtDes_Present 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
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
         Left            =   1185
         MaxLength       =   50
         TabIndex        =   8
         Top             =   645
         Width           =   4365
      End
      Begin FunctionsButtons.FunctButt funPresent 
         Height          =   510
         Left            =   1620
         TabIndex        =   9
         Top             =   1350
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmTG_ColCli.frx":0000
         Orientacion     =   0
         Style           =   1
         Language        =   1
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
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
         Index           =   7
         Left            =   105
         TabIndex        =   19
         Tag             =   "Id Color"
         Top             =   720
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdCod_EstCli 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4440
      TabIndex        =   16
      Top             =   855
      Width           =   405
   End
   Begin VB.TextBox txtCod_ColCliPre 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
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
      Left            =   1395
      MaxLength       =   20
      TabIndex        =   0
      Top             =   30
      Width           =   2310
   End
   Begin VB.TextBox txtNom_ColCliPre 
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
      Left            =   1380
      MaxLength       =   30
      TabIndex        =   1
      Top             =   420
      Width           =   4335
   End
   Begin VB.Frame fraAddColor 
      Caption         =   "Creación de Colores"
      Height          =   2160
      Left            =   75
      TabIndex        =   10
      Top             =   2100
      Visible         =   0   'False
      Width           =   5745
      Begin VB.TextBox txtCod_ColCli 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
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
         Left            =   1275
         MaxLength       =   20
         TabIndex        =   4
         Top             =   240
         Width           =   2310
      End
      Begin VB.TextBox txtNom_ColCli 
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
         Left            =   1275
         MaxLength       =   30
         TabIndex        =   5
         Top             =   615
         Width           =   4335
      End
      Begin FunctionsButtons.FunctButt acbForm 
         Height          =   510
         Left            =   1620
         TabIndex        =   7
         Top             =   1530
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmTG_ColCli.frx":0097
         Orientacion     =   0
         Style           =   1
         Language        =   1
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin SSDataWidgets_B.SSDBCombo cboTipColor 
         Height          =   300
         Left            =   1275
         TabIndex        =   6
         Top             =   1005
         Width           =   2340
         DataFieldList   =   "Column 0"
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   953
         Columns(0).Caption=   "Tipo"
         Columns(0).Name =   "Cod_TipCol"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3175
         Columns(1).Caption=   "Descripción"
         Columns(1).Name =   "Nom_TipCol"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   4128
         _ExtentY        =   529
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Column 1"
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
         Index           =   2
         Left            =   90
         TabIndex        =   13
         Tag             =   "Description"
         Top             =   675
         Width           =   945
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Id Color :"
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
         Left            =   90
         TabIndex        =   12
         Tag             =   "Id Color"
         Top             =   300
         Width           =   630
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Color :"
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
         TabIndex        =   11
         Tag             =   "Color Type"
         Top             =   1050
         Width           =   1035
      End
   End
   Begin FunctionsButtons.FunctButt funAceptaCancelar 
      Height          =   510
      Left            =   1665
      TabIndex        =   3
      Top             =   1425
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmTG_ColCli.frx":012E
      Orientacion     =   0
      Style           =   1
      Language        =   1
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin SSDataWidgets_B.SSDBCombo cboCod_Present 
      Height          =   315
      Left            =   1380
      TabIndex        =   2
      Top             =   840
      Width           =   2955
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   953
      Columns(0).Caption=   "Clase"
      Columns(0).Name =   "Cod_Present"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3175
      Columns(1).Caption=   "Descripción"
      Columns(1).Name =   "Des_Present"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   5212
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 1"
   End
   Begin VB.Label Etiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Presentaciones :"
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
      Index           =   3
      Left            =   210
      TabIndex        =   17
      Tag             =   "Description"
      Top             =   885
      Width           =   1215
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
      Index           =   5
      Left            =   195
      TabIndex        =   15
      Tag             =   "Description"
      Top             =   465
      Width           =   945
   End
   Begin VB.Label Etiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Id Color :"
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
      Index           =   4
      Left            =   195
      TabIndex        =   14
      Tag             =   "Id Color"
      Top             =   90
      Width           =   630
   End
End
Attribute VB_Name = "frmTG_ColCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sCod_Cliente As String

Public bOk          As Boolean

Public scod_colcli  As String

Public sCod_TemCli  As String

Public sCod_EstCli  As String

Public sCod_EstPro  As String

Public sFlag        As String

Public sModalAdd    As String

Private Sub acbForm_ActionClick(ByVal Index As Integer, _
                                ByVal ActionType As Integer, _
                                ByVal ActionName As String)

    On Error GoTo errores

    Dim vbuff

    Dim obj       As clsTG_ColCli

    Dim sTipColor As String
    
    Select Case ActionName

        Case "ACEPTAR"

            If txtCod_ColCli.Text = "" Then
                If txtCod_ColCli.Enabled Then
                    txtCod_ColCli.SetFocus
                End If

                Exit Sub

            End If
        
            If txtNom_ColCli.Text = "" Then
                If txtNom_ColCli.Enabled Then
                    txtNom_ColCli.SetFocus
                End If

                Exit Sub

            End If
        
            If cboTipColor.value = "" Then
                If cboTipColor.Enabled Then
                    cboTipColor.SetFocus
                End If
            End If
        
            sTipColor = cboTipColor.value
        
            Set obj = New clsTG_ColCli
            obj.ConexionString = cCONNECT
            obj.Add sCod_Cliente, txtCod_ColCli.Text, txtNom_ColCli.Text, sTipColor
            Set obj = Nothing
        
            Me.txtCod_ColCliPre = Me.txtCod_ColCli.Text
            Me.txtNom_ColCliPre = Me.txtNom_ColCli.Text
            'Me.txtNom_ColCliPre.Enabled = False
            Me.cboCod_Present.SetFocus
            Me.cboCod_Present.DroppedDown = True
        
            sModalAdd = "ADICIONAR"
            Me.fraAddColor.Visible = False
        
        Case "CANCELAR"
            Me.fraAddColor.Visible = False
    End Select

    Exit Sub

errores:

    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description
End Sub

Private Sub cmdCod_EstCli_Click()
    Me.fraAddPresent.Visible = True
    Me.txtDes_Present.Text = txtNom_ColCliPre.Text
    Me.txtDes_Present.SetFocus
End Sub

Private Sub Form_Load()
    CargarCboColores
    
End Sub

Private Sub CargarCboColores()

    On Error GoTo errores

    Dim vbuff

    Dim obj As New clsTG_ColCli

    Dim i   As Long
           
    Set obj = New clsTG_ColCli
    obj.ConexionString = cCONNECT
    vbuff = obj.ViewTipCol
    
    cboTipColor.TagVariant = cboTipColor.Cols
    cboTipColor.RemoveAll
    LibraryVBToSSDBCombo obj, vbuff, cboTipColor
    Set obj = Nothing
    
    Exit Sub

errores:

    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description
End Sub

Public Sub LibraryVBToSSDBCombo(ByRef oData As Object, _
                                ByRef pBuff As Variant, _
                                ByRef ssDBCombo As SSDataWidgets_B.ssDBCombo)

    On Error Resume Next

    Dim rsBuff    As LibraryVB.clsRecords

    Dim iContador As Long

    Dim nCols     As Integer

    Dim iVerif    As Integer

    Dim temp      As String

    Dim NVEZ      As Boolean

    Dim X%

    Dim total1    As Long

    Dim y%

    Dim i         As Long

    Dim ic        As Long

    Dim bPrimero  As Boolean

    ssDBCombo.FieldSeparator = "~"
    Set rsBuff = New LibraryVB.clsRecords
    Set rsBuff.RefObject = oData

    rsBuff.Buffer = pBuff
    ssDBCombo.Redraw = False
    nCols = rsBuff.count

    ic = ssDBCombo.Cols

    If ssDBCombo.Cols < nCols Then

        For i = nCols To ic + 1 Step -1
            ssDBCombo.Columns.Add ssDBCombo.Cols    ' "Column" & i, 500, False, Nothing, "Column" & i
            ssDBCombo.Columns(ssDBCombo.Cols - 1).Name = rsBuff(ssDBCombo.Cols).Name
            ssDBCombo.Columns(ssDBCombo.Cols - 1).Caption = rsBuff(ssDBCombo.Cols).Name
        Next i

    End If

    For y = 0 To ssDBCombo.Cols - 1

        If ssDBCombo.Columns(y).DataType = 5 Or ssDBCombo.Columns(y).DataType = 6 Or ssDBCombo.Columns(y).DataType = 9 Then
            ssDBCombo.Columns(y).TagVariant = 0
        End If

    Next

    NVEZ = True

    bPrimero = True
    X = 0
    ssDBCombo.RemoveAll

    Do While Not rsBuff.EOF
        temp = ""

        For iContador = 0 To nCols - 1
            ssDBCombo.Columns(iContador).Locked = True
            ssDBCombo.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
            ssDBCombo.Columns(iContador).Style = 4 'ssStyleButton
            temp = temp & FixNulos(rsBuff(iContador + 1), vbstring)
            
            If iContador < nCols - 1 Then
                temp = temp & "~"
            End If

            If iContador >= FixNulos(ssDBCombo.TagVariant, vbLong) Then
                ssDBCombo.Columns(iContador).DataType = 5
                ssDBCombo.Columns(iContador).Alignment = 1
            End If

            'ssdbCombo.Columns(iContador).DataType = 5
            If ssDBCombo.Columns(iContador).DataType = 5 Or ssDBCombo.Columns(iContador).DataType = 6 Or ssDBCombo.Columns(iContador).DataType = 9 Or iContador > FixNulos(ssDBCombo.TagVariant, vbLong) Then
                If Val(FixNulos(rsBuff(iContador + 1), vbDouble)) > 0 Then
                    ssDBCombo.Columns(iContador).TagVariant = Val(ssDBCombo.Columns(iContador).TagVariant) + FixNulos(rsBuff(iContador + 1), vbDouble)
                End If
            End If

        Next

        NVEZ = False
        ssDBCombo.AddItem temp
        rsBuff.MoveNext
        X = X + 1
    Loop
 
    ssDBCombo.RowHeight = 300 ' ssdbCombo.RowHeight * 1.25
    ssDBCombo.Refresh

    ssDBCombo.Redraw = True
    Set rsBuff.RefObject = Nothing
    Set rsBuff = Nothing

End Sub

Private Sub funAceptaCancelar_ActionClick(ByVal Index As Integer, _
                                          ByVal ActionType As Integer, _
                                          ByVal ActionName As String)

    On Error GoTo errores

    Dim obj As clsTG_ColCli
    
    Select Case ActionName

        Case "ACEPTAR"

            If Me.txtCod_ColCliPre.Text = "" Then
                Mensaje kMESSAGE_ERR_NOTEMPTY

                If Me.txtCod_ColCliPre.Enabled Then
                    Me.txtCod_ColCliPre.SetFocus
                End If

                Exit Sub

            End If
        
            If Me.cboCod_Present.Text = "" Then
                Mensaje kMESSAGE_ERR_NOTEMPTY

                If Me.cboCod_Present.Enabled Then
                    Me.cboCod_Present.SetFocus
                End If

                Exit Sub

            End If
        
            sFlag = "COD_COLCLIPRE2"
            scod_colcli = Me.txtCod_ColCliPre.Text

            If Filtrar(sFlag, Me, txtCod_ColCliPre, txtNom_ColCliPre, False) Then
                If Not ValidarColCli(sCod_Cliente, sCod_TemCli, sCod_EstCli, scod_colcli) Then
                    Mensaje kMESSAGE_ERR_CODIGO_YA_REGISTRADO

                    If Me.txtCod_ColCliPre.Enabled Then
                        Me.txtCod_ColCliPre.SetFocus
                    End If

                    Exit Sub

                End If

                Me.cboCod_Present.SetFocus
            Else
        
            End If
        
            Set obj = New clsTG_ColCli
            obj.ConexionString = cCONNECT
            obj.AddEStCliCol sCod_Cliente, Me.sCod_TemCli, Me.sCod_EstCli, Me.txtCod_ColCliPre.Text, Me.txtNom_ColCliPre.Text, Me.sCod_EstPro, Me.cboCod_Present.value, vusu & " " & ComputerName()
            Set obj = Nothing
        
            scod_colcli = Me.txtCod_ColCliPre.Text
            bOk = True
            Unload Me

        Case "CANCELAR"
            Unload Me
    End Select

    Exit Sub

errores:

    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description
    
End Sub

Private Sub funPresent_ActionClick(ByVal Index As Integer, _
                                   ByVal ActionType As Integer, _
                                   ByVal ActionName As String)

    Select Case ActionName

        Case "ACEPTAR"
            txtDes_Present.Text = UCase(txtDes_Present.Text)

            If AddPresent(sCod_EstPro, Trim(txtDes_Present.Text)) Then
                CargarPresentaciones sCod_EstPro
            End If

            Me.fraAddPresent.Visible = False

        Case "CANCELAR"
            Me.fraAddPresent.Visible = False
    End Select

End Sub

Private Sub txtCod_ColCliPre_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If Me.txtNom_ColCliPre.Enabled Then
            Me.txtNom_ColCliPre.Text = Me.txtCod_ColCliPre.Text
            Me.txtNom_ColCliPre.SetFocus
        End If
    End If

End Sub

Private Function ValidarColCli(sCod_Cliente, _
                               sCod_TemCli, _
                               sCod_EstCli, _
                               scod_colcli) As Boolean

    On Error GoTo errores

    Dim vbuff

    Dim obj As New clsTG_ColCli

    Set obj = New clsTG_ColCli
    obj.ConexionString = cCONNECT
    vbuff = obj.ViewTG_EstCliCol(sCod_Cliente, sCod_TemCli, sCod_EstCli, scod_colcli)
    Set obj = Nothing

    If IsEmpty(vbuff) Then
        ValidarColCli = True
    Else
        ValidarColCli = False
    End If

    Exit Function

errores:

    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description

End Function

Private Function AddPresent(ByVal sCod_EstPro As String, _
                            ByVal sDes_Present As String) As Boolean

    On Error GoTo errores

    Dim obj As New clsTG_ColCli
                
    Set obj = New clsTG_ColCli
    obj.ConexionString = cCONNECT
    obj.AddEstProPre sCod_EstPro, sDes_Present
    Set obj = Nothing
        
    AddPresent = True

    Exit Function

errores:

    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description
End Function

Public Sub CargarPresentaciones(ByVal sCod_EstPro As String)

    On Error GoTo errores

    Dim vbuff

    Dim obj As New clsTG_ColCli

    Dim i   As Long
           
    Set obj = New clsTG_ColCli
    obj.ConexionString = cCONNECT
    vbuff = obj.ViewAllPresent_EstPro(sCod_EstPro)
    
    cboCod_Present.TagVariant = cboCod_Present.Cols
    cboCod_Present.RemoveAll
    LibraryVBToSSDBCombo obj, vbuff, cboCod_Present
    Set obj = Nothing
    
    Exit Sub

errores:

    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description

End Sub

Private Sub txtCod_ColCliPre_LostFocus()
    Me.txtNom_ColCliPre.Text = txtCod_ColCliPre.Text
    Me.txtNom_ColCliPre.SetFocus
End Sub

