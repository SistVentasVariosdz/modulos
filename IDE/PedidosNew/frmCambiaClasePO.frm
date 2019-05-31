VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form frmCambiaClasePO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar clase PO."
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   0
      TabIndex        =   8
      Top             =   1530
      Width           =   4905
      Begin SSDataWidgets_B.SSDBCombo cbo_ClasePO 
         Height          =   285
         Left            =   1575
         TabIndex        =   9
         Top             =   240
         Width           =   2625
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
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Caption=   "Clase"
         Columns(0).Name =   "Cod_ClaPurord"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "Descripcion"
         Columns(1).Name =   "Des_Clapurord"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   4630
         _ExtentY        =   503
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Column 1"
      End
      Begin VB.Label Label1 
         Caption         =   "Nueva clase PO"
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
         Left            =   75
         TabIndex        =   10
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.Frame fraDatos 
      Height          =   1500
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4920
      Begin VB.TextBox txtPO 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   4
         Top             =   285
         Width           =   2940
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   3
         Top             =   645
         Width           =   2940
      End
      Begin VB.TextBox txtXClasePO 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1005
         Width           =   2955
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "PO :"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Tag             =   "PO :"
         Top             =   345
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Tag             =   "Client :"
         Top             =   705
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Clase PO :"
         Height          =   195
         Left            =   225
         TabIndex        =   5
         Tag             =   "Style :"
         Top             =   1035
         Width           =   750
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1080
      TabIndex        =   0
      Top             =   2430
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmCambiaClasePO.frx":0000
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmCambiaClasePO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public varCod_Cliente, varNom_Cliente, varCod_Purord, varCod_Clapurord As String

Dim mRs As New Recordset

Private Sub Cargar_ClasePO()

    On Error GoTo errores

    Dim vbuff

    Dim obj As New clsTG_ColCli

    Dim i   As Long
           
    Set obj = New clsTG_ColCli
    obj.ConexionString = cCONNECT
    vbuff = obj.ViewClasePO(varCod_Clapurord)
    
    cbo_ClasePO.TagVariant = cbo_ClasePO.Cols
    cbo_ClasePO.RemoveAll
    LibraryVBToSSDBCombo obj, vbuff, cbo_ClasePO
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

Private Sub Form_Load()
    Call Cargar_ClasePO
    txtCliente.Text = varNom_Cliente
    txtPO.Text = varCod_Purord
    txtXClasePO.Text = varCod_Clapurord
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, _
                                   ByVal ActionType As Integer, _
                                   ByVal ActionName As String)

    Select Case ActionName

        Case "ACEPTAR"

            If Trim(cbo_ClasePO.Text) = "" Then
                MsgBox "Ingrese el clase de PO", vbInformation, "Aviso"
                cbo_ClasePO.SetFocus

                Exit Sub

            End If

            On Error GoTo errsalvar

            strSql = "EXEC tg_cambia_clapurord_po '" & varCod_Cliente & "','" & varCod_Purord & "','" & cbo_ClasePO.Columns(0).value & "'"
            Call ExecuteCommandSQL(cCONNECT, strSql)
            Call MsgBox("Funcion correctamente realizada. Sirvase verificar", vbInformation)
            Unload Me

            Exit Sub

errsalvar:
            ErrorHandler Err, "SALVAR_DATOS"
            Unload Me

        Case "CANCELAR"
            Unload Me
    End Select

End Sub
