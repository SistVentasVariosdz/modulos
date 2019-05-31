VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmShowTG_PurOrd_EstCli_Grafico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ASIGNAR GRÁFICOS"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7545
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9105
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8640
         TabIndex        =   13
         Top             =   1800
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   435
         Left            =   8010
         TabIndex        =   4
         Top             =   7050
         Width           =   1065
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   435
         Left            =   6900
         TabIndex        =   3
         Top             =   7050
         Width           =   1065
      End
      Begin VB.TextBox txtGrafico 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   0
         Top             =   960
         Width           =   1785
      End
      Begin VB.TextBox txtEstiloCliente 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         Top             =   510
         Width           =   3945
      End
      Begin VB.TextBox txtPO 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         Top             =   150
         Width           =   3945
      End
      Begin GridEX20.GridEX GEXListaGrafico 
         Height          =   5175
         Left            =   60
         TabIndex        =   1
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   9128
         Version         =   "2.0"
         RecordNavigator =   -1  'True
         RecordNavigatorString=   ""
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         GridLineStyle   =   2
         RowHeight       =   30
         MethodHoldFields=   -1  'True
         GroupByBoxVisible=   0   'False
         BackColorHeader =   -2147483626
         HeaderFontName  =   "Verdana"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   6.75
         HeaderFontWeight=   700
         GridLines       =   2
         ColumnHeaderHeight=   465
         IntProp1        =   0
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmShowTG_PurOrd_EstCli_Grafico.frx":0000
         FormatStyle(2)  =   "frmShowTG_PurOrd_EstCli_Grafico.frx":0128
         FormatStyle(3)  =   "frmShowTG_PurOrd_EstCli_Grafico.frx":01D8
         FormatStyle(4)  =   "frmShowTG_PurOrd_EstCli_Grafico.frx":028C
         FormatStyle(5)  =   "frmShowTG_PurOrd_EstCli_Grafico.frx":0364
         FormatStyle(6)  =   "frmShowTG_PurOrd_EstCli_Grafico.frx":041C
         FormatStyle(7)  =   "frmShowTG_PurOrd_EstCli_Grafico.frx":04FC
         ImageCount      =   0
         PrinterProperties=   "frmShowTG_PurOrd_EstCli_Grafico.frx":0608
      End
      Begin GridEX20.GridEX GEXListaColores 
         Height          =   5175
         Left            =   2250
         TabIndex        =   2
         Top             =   1680
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   9128
         Version         =   "2.0"
         RecordNavigator =   -1  'True
         RecordNavigatorString=   "Registros:|de"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         GridLineStyle   =   2
         RowHeight       =   30
         MethodHoldFields=   -1  'True
         GroupByBoxVisible=   0   'False
         BackColorHeader =   -2147483626
         HeaderFontName  =   "Verdana"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   6.75
         HeaderFontWeight=   700
         GridLines       =   2
         ColumnHeaderHeight=   465
         IntProp1        =   0
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmShowTG_PurOrd_EstCli_Grafico.frx":07E0
         FormatStyle(2)  =   "frmShowTG_PurOrd_EstCli_Grafico.frx":0908
         FormatStyle(3)  =   "frmShowTG_PurOrd_EstCli_Grafico.frx":09B8
         FormatStyle(4)  =   "frmShowTG_PurOrd_EstCli_Grafico.frx":0A6C
         FormatStyle(5)  =   "frmShowTG_PurOrd_EstCli_Grafico.frx":0B44
         FormatStyle(6)  =   "frmShowTG_PurOrd_EstCli_Grafico.frx":0BFC
         FormatStyle(7)  =   "frmShowTG_PurOrd_EstCli_Grafico.frx":0CDC
         ImageCount      =   0
         PrinterProperties=   "frmShowTG_PurOrd_EstCli_Grafico.frx":0DE8
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00CCCCCC&
         Caption         =   "LISTA DE GRÁFICOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1545
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00CCCCCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00666666&
         Height          =   315
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   1380
         Width           =   2205
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "GRÁFICO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1020
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ESTILO CLIENTE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   10
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   210
         Width           =   210
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00CCCCCC&
         Caption         =   "LISTA DE COLORES CLIENTE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2310
         TabIndex        =   7
         Top             =   1440
         Width           =   2130
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00666666&
         Height          =   6915
         Left            =   30
         Top             =   30
         Width           =   9045
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00CCCCCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00666666&
         Height          =   315
         Left            =   2250
         Shape           =   4  'Rounded Rectangle
         Top             =   1380
         Width           =   6765
      End
   End
End
Attribute VB_Name = "frmShowTG_PurOrd_EstCli_Grafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public strCodCliente As String

Private strSql       As String

Private boolGuardo   As Boolean

Private Sub chkTodos_Click()

    Dim i      As Integer

    Dim rsTemp As New ADODB.Recordset

    GEXListaColores.Update
    Set rsTemp = GEXListaColores.ADORecordset

    If rsTemp.RecordCount Then

        Do While Not rsTemp.EOF
        
            rsTemp.Fields("SEL").value = chkTodos.value
            rsTemp.MoveNext
        Loop

        Set GEXListaColores.ADORecordset = rsTemp
        Call FORMATO_GRILLA
    End If

End Sub

Private Sub checkedPadre()

    Dim i           As Integer

    Dim boolChecked As Boolean

    boolChecked = True

    Dim rsTemp As New ADODB.Recordset

    GEXListaColores.Update

    Set rsTemp = GEXListaColores.ADORecordset

    If rsTemp.RecordCount Then

        Do While Not rsTemp.EOF
            
            If rsTemp.Fields("SEL").value = False Then
                boolChecked = False
                chkTodos.value = 0
            End If

            rsTemp.MoveNext
        Loop

        If boolChecked Then
            chkTodos.value = 1
        End If
    End If

End Sub

Private Sub cmdAceptar_Click()

    Dim i As Integer

    GEXListaColores.Update

    If (GEXListaColores.RowCount > 0) Then
        GEXListaColores.MoveFirst

        For i = 1 To GEXListaColores.RowCount
            Call TRANSAC_DATOS_GRAFICO_COLOR("I")
            GEXListaColores.MoveNext
        Next

    End If

    Call TRANSAC_DATOS_GRAFICO_COLOR("C")
    MsgBox "Los colores y gráficos fueron actualizados correctamente", vbOKOnly + vbInformation, Me.Caption
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub GEXListaColores_DblClick()
    Call TRANSAC_DATOS_GRAFICO_COLOR("D")
    Call TRANSAC_DATOS_GRAFICO_COLOR("C")
End Sub

Private Sub GEXListaGrafico_Click()
    Call TRANSAC_DATOS_GRAFICO_COLOR("C")
    
End Sub

Private Sub GEXListaGrafico_DblClick()
    Call TRANSAC_DATOS_GRAFICO("D")
    Call TRANSAC_DATOS_GRAFICO("C")
End Sub

Private Sub txtGrafico_GotFocus()
    Call SelectionText(txtGrafico)
End Sub

Private Sub txtGrafico_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call TRANSAC_DATOS_GRAFICO("I")
        Call TRANSAC_DATOS_GRAFICO("C")
        txtGrafico.Text = ""
        txtGrafico.SetFocus
        
    Else
        Call SoloNumeros(txtGrafico, KeyAscii, False, 0, 6)
    End If

End Sub

'TG_PurOrd_EstCli_Graficos_MAN
'@OPCION
'@Cod_PurOrd
'@Cod_EstCli
'@Cod_Grafico

Public Sub TRANSAC_DATOS_GRAFICO(strTipo As String)
    
    On Error GoTo Salvar_DatosErr
    
    boolGuardo = False

    If (strTipo = "I") Then

        strSql = "EXEC TG_PurOrd_EstCli_Graficos_MAN " & vbNewLine
        strSql = strSql & "@OPCION = '" & strTipo & "'" & vbNewLine
        strSql = strSql & ",@Cod_PurOrd = '" & Trim(txtPO) & "'" & vbNewLine
        strSql = strSql & ",@Cod_EstCli = '" & Trim(txtEstiloCliente) & "'" & vbNewLine
        strSql = strSql & ",@Cod_Grafico = '" & Trim(txtGrafico) & "'" & vbNewLine
        ExecuteCommandSQL cCONNECT, strSql
        
    ElseIf (strTipo = "C") Then

        strSql = "EXEC TG_PurOrd_EstCli_Graficos_MAN " & vbNewLine
        strSql = strSql & "@OPCION = '" & strTipo & "'" & vbNewLine
        strSql = strSql & ",@Cod_PurOrd = '" & Trim(txtPO) & "'" & vbNewLine
        strSql = strSql & ",@Cod_EstCli = '" & Trim(txtEstiloCliente) & "'" & vbNewLine
        strSql = strSql & ",@Cod_Grafico = '" & Trim(txtGrafico) & "'" & vbNewLine
        Set GEXListaGrafico.ADORecordset = CargarRecordSetDesconectado(strSql, cCONNECT)
        
        With GEXListaGrafico

            For n = 1 To .Columns.count
                .Columns.ItemByPosition(n).EditType = jgexEditNone
                .Columns.ItemByPosition(n).HeaderAlignment = jgexAlignCenter
                .Columns.ItemByPosition(n).Caption = UCase(.Columns.ItemByPosition(n).Caption)
                .Columns.ItemByPosition(n).WordWrap = True
                .Columns.ItemByPosition(n).Visible = False
                .Columns.ItemByPosition(n).Width = 1500
            Next

        End With

        With GEXListaGrafico.Columns("Cod_Grafico")
            .Caption = "GRÁFICO"
            .Visible = True
            .Width = 1800
            .TextAlignment = jgexAlignCenter
            .ColPosition = 1
        End With

    ElseIf (strTipo = "D") Then
        strSql = "EXEC TG_PurOrd_EstCli_Graficos_MAN " & vbNewLine
        strSql = strSql & "@OPCION = '" & strTipo & "'" & vbNewLine
        strSql = strSql & ",@Cod_PurOrd = '" & Trim(txtPO) & "'" & vbNewLine
        strSql = strSql & ",@Cod_EstCli = '" & Trim(txtEstiloCliente) & "'" & vbNewLine
        strSql = strSql & ",@Cod_Grafico = '" & Trim(GEXListaGrafico.value(GEXListaGrafico.Columns("Cod_Grafico").Index)) & "'" & vbNewLine
        ExecuteCommandSQL cCONNECT, strSql
    End If

    boolGuardo = True

    Exit Sub

Salvar_DatosErr:
    MsgBox Err.Description, vbExclamation, Me.Caption:
    guardo = False
End Sub

Public Sub TRANSAC_DATOS_GRAFICO_COLOR(strTipo As String)
    
    On Error GoTo Salvar_DatosErr
    
    boolGuardo = False

    If (strTipo = "I") Then

        strSql = "EXEC TG_PurOrd_EstCli_Graficos_Colores_MAN " & vbNewLine
        strSql = strSql & "@OPCION = '" & strTipo & "'" & vbNewLine
        strSql = strSql & ",@Cod_PurOrd = '" & Trim(txtPO) & "'" & vbNewLine
        strSql = strSql & ",@Cod_EstCli = '" & Trim(txtEstiloCliente) & "'" & vbNewLine
        strSql = strSql & ",@Cod_Grafico = '" & Trim(Trim(GEXListaGrafico.value(GEXListaGrafico.Columns("Cod_Grafico").Index))) & "'" & vbNewLine
        strSql = strSql & ",@Cod_Cliente = '" & strCodCliente & "'" & vbNewLine
        strSql = strSql & ",@Cod_ColCli = '" & Trim(GEXListaColores.value(GEXListaColores.Columns("Cod_ColCli").Index)) & "'" & vbNewLine
        strSql = strSql & ",@Flg_Selec = " & IIf(GEXListaColores.value(GEXListaColores.Columns("SEL").Index), 1, 0) & "" & vbNewLine
        ExecuteCommandSQL cCONNECT, strSql
        
    ElseIf (strTipo = "C") Then

        strSql = "EXEC TG_PurOrd_EstCli_Graficos_Colores_MAN " & vbNewLine
        strSql = strSql & "@OPCION = '" & strTipo & "'" & vbNewLine
        strSql = strSql & ",@Cod_PurOrd = '" & Trim(txtPO) & "'" & vbNewLine
        strSql = strSql & ",@Cod_EstCli = '" & Trim(txtEstiloCliente) & "'" & vbNewLine
        strSql = strSql & ",@Cod_Grafico = '" & Trim(Trim(GEXListaGrafico.value(GEXListaGrafico.Columns("Cod_Grafico").Index))) & "'" & vbNewLine
        strSql = strSql & ",@Cod_Cliente = '" & strCodCliente & "'" & vbNewLine
        strSql = strSql & ",@Cod_ColCli = '" & Trim("") & "'" & vbNewLine
        strSql = strSql & ",@Flg_Selec = " & 0 & "" & vbNewLine
        Set GEXListaColores.ADORecordset = CargarRecordSetDesconectado(strSql, cCONNECT)
     
        chkTodos.Visible = True
        chkTodos.Left = 8340
        checkedPadre
        Call FORMATO_GRILLA

    ElseIf (strTipo = "D") Then
        strSql = "EXEC TG_PurOrd_EstCli_Graficos_Colores_MAN " & vbNewLine
        strSql = strSql & "@OPCION = '" & strTipo & "'" & vbNewLine
        strSql = strSql & ",@Cod_PurOrd = '" & Trim(txtPO) & "'" & vbNewLine
        strSql = strSql & ",@Cod_EstCli = '" & Trim(txtEstiloCliente) & "'" & vbNewLine
        strSql = strSql & ",@Cod_Cliente = '" & strCodCliente & "'" & vbNewLine
        strSql = strSql & ",@Cod_Grafico = '" & Trim(Trim(GEXListaGrafico.value(GEXListaGrafico.Columns("Cod_Grafico").Index))) & "'" & vbNewLine
        strSql = strSql & ",@Cod_ColCli = '" & Trim(GEXListaColores.value(GEXListaColores.Columns("Cod_ColCli").Index)) & "'" & vbNewLine
        strSql = strSql & ",@Flg_Selec = " & IIf(GEXListaColores.value(GEXListaColores.Columns("SEL").Index), 1, 0) & "" & vbNewLine
        ExecuteCommandSQL cCONNECT, strSql
    End If

    boolGuardo = True

    Exit Sub

Salvar_DatosErr:
    MsgBox Err.Description, vbExclamation, Me.Caption:
    guardo = False
End Sub

Private Sub FORMATO_GRILLA()

    With GEXListaColores

        For n = 1 To .Columns.count
            .Columns.ItemByPosition(n).EditType = jgexEditNone
            .Columns.ItemByPosition(n).HeaderAlignment = jgexAlignCenter
            .Columns.ItemByPosition(n).Caption = UCase(.Columns.ItemByPosition(n).Caption)
            .Columns.ItemByPosition(n).WordWrap = True
            .Columns.ItemByPosition(n).Visible = False
            .Columns.ItemByPosition(n).Width = 1300
        Next

    End With
   
    With GEXListaColores.Columns("Cod_ColCli")
        .Caption = "COD. COLOR"
        .Visible = True
        .Width = 1500
        .TextAlignment = jgexAlignCenter
        .ColPosition = 1
    End With

    With GEXListaColores.Columns("Nom_ColCli")
        .Caption = "NOM. COLOR"
        .Visible = True
        .Width = 2200
        .TextAlignment = jgexAlignCenter
        .ColPosition = 2
    End With

    With GEXListaColores.Columns("Color_Segun_Cliente")
        .Caption = "NOM. COLOR CLIENTE"
        .Visible = True
        .Width = 2200
        .TextAlignment = jgexAlignCenter
        .ColPosition = 3
    End With

    With GEXListaColores.Columns("SEL")
        .Caption = ""
        .Visible = True
        .Width = 500
        .TextAlignment = jgexAlignCenter
        .CellStyle = jgexCheckBox
        .EditType = jgexEditCheckBox
        .ColumnType = jgexCheckBox
        .ColPosition = 4
            
    End With

End Sub
