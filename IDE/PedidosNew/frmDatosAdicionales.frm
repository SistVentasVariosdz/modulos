VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDatosAdicionales 
   Caption         =   "Asignar Nro Despacho -PackOne"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_purchar 
      Height          =   285
      Left            =   2520
      TabIndex        =   11
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txt4 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin FunctionsButtons.FunctButt funTemCli 
      Height          =   510
      Left            =   960
      TabIndex        =   0
      Top             =   2280
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmDatosAdicionales.frx":0000
      Orientacion     =   0
      Style           =   1
      Language        =   1
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   315
      Left            =   2520
      TabIndex        =   6
      Top             =   840
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "dd/mm/yyyy"
      Format          =   58130433
      CurrentDate     =   39164.7102662037
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   2520
      TabIndex        =   7
      Top             =   480
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "dd/mm/yyyy"
      Format          =   58130433
      CurrentDate     =   39164.7101388889
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2520
      TabIndex        =   8
      Top             =   120
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "dd/mm/yyyy"
      Format          =   58130433
      CurrentDate     =   39164.7099537037
   End
   Begin VB.Label Label6 
      Caption         =   "Purchase Order Agente"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Num_Booking"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   2205
   End
   Begin VB.Label Label4 
      Caption         =   "Num_Booking"
      Height          =   270
      Left            =   -2040
      TabIndex        =   5
      Top             =   -1200
      Width           =   2205
   End
   Begin VB.Label Label3 
      Caption         =   "Actual ExFactory Date"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2205
   End
   Begin VB.Label Label2 
      Caption         =   "Recepcion UPC"
      Height          =   270
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2085
   End
   Begin VB.Label Label1 
      Caption         =   "ExFactoryDate Confirmed"
      Height          =   270
      Left            =   255
      TabIndex        =   1
      Top             =   195
      Width           =   1965
   End
End
Attribute VB_Name = "frmDatosAdicionales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sCod_Cliente     As String

Public sCod_PurOrd      As String

Public sCod_LotPurOrd   As String

Public sCod_EstCli      As String

Public Codigo           As String

Public Descripcion      As String

Public sAccionName      As String

Public sModoWizard      As String

Public sCod_TemCli      As String

Public oParent          As Object

Dim strSql              As String

Public sFlgOpcion_Nueva As String

Public Function BUSCAR() As Boolean

    On Error Resume Next

    'On Error GoTo errores
    Dim vbuff

    Dim objPO     As clsTG_LotColTal

    Dim rsBuff    As LibraryVB.clsRecords

    Dim varStrsql As String

    Dim i         As Integer
    
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
        
    Set rsBuff = New LibraryVB.clsRecords
    Set rsBuff.RefObject = objPO
           
    rsBuff.Buffer = objPO.LoadDataPOC_mp(sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli)
        
    Do While Not rsBuff.EOF
        Me.sCod_PurOrd = FixNulos(rsBuff!cod_purord, vbstring)
        'BuscarComboD cboCod_ClaPurOrd, FixNulos(rsBuff!Cod_ClaPurOrd, vbString)
        'cboCod_ClaPurOrd.value = FixNulos(rsBuff!Cod_ClaPurOrd, vbstring)
        'cboCod_ClaPurOrd.Enabled = False
           
        Me.DTPicker1.value = FixNulos(rsBuff!Fec_ExFactoryDate_Confirmed, vbstring)
        Me.DTPicker2.value = FixNulos(rsBuff!Fec_Recepcion_UPC, vbstring)
        Me.DTPicker3.value = FixNulos(rsBuff!Fec_ActualExFactoryDate, vbstring)
        Me.txt4.Text = FixNulos(rsBuff!Num_Booking, vbstring)
        Me.txt4.Text = Trim(Me.txt4.Text)
        Me.txt_purchar.Text = FixNulos(rsBuff!Cod_PurOrd_Agente, vbstring)
        rsBuff.MoveNext
    Loop
    
    Set rsBuff.RefObject = Nothing
    Set rsBuff = Nothing
    Set objPO = Nothing
End Function

Public Sub CargaNroDespachoActual()

    On Error GoTo errx

    Dim sSQl As String

    sSQl = "EXEC SM_CONSULTA_NRO_DESPACHO_LOTEST '$','$','$','$'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli)

    Exit Sub

errx:
    Err.Raise Err.Number, Err.source, Err.Description
End Sub

Private Sub Form_Load()
    funTemCli.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    'Me.txt_purchar.Text = sCod_PurOrd
End Sub

Private Sub funTemCli_ActionClick(ByVal Index As Integer, _
                                  ByVal ActionType As Integer, _
                                  ByVal ActionName As String)

    Select Case ActionName

        Case "ACEPTAR"
            Grabar

        Case "SALIR"
            Unload Me
    End Select

End Sub

Private Sub Grabar()

    On Error GoTo errx

    Dim sSQl    As String

    Dim dfecha1 As String

    Dim dfecha2 As String

    Dim dfecha3 As String

    If IsNull(DTPicker1.value) Then
        DTPicker1.value = ""
    Else
        DTPicker1.value = DTPicker1.value
    End If

    If IsNull(DTPicker2.value) Then
        DTPicker2.value = ""
    Else
        DTPicker2.value = DTPicker2.value
    End If

    If IsNull(DTPicker3.value) Then
        DTPicker3.value = ""
    Else
        DTPicker3.value = DTPicker3.value
    End If

    dfecha1 = Format(DTPicker1.value, "dd/mm/yyyy")
    dfecha2 = Format(DTPicker2.value, "dd/mm/yyyy")
    dfecha3 = Format(DTPicker3.value, "dd/mm/yyyy")

    'sSql = "EXEC SM_DATOS_ADICIONALES '$','$','$','$','$','$','$','$','$'"
    sSQl = "EXEC SM_DATOS_ADICIONALES 'U','" & Trim(sCod_Cliente) & "','" & Trim(sCod_PurOrd) & "','" & Trim(sCod_LotPurOrd) & "','" & Trim(sCod_EstCli) & "','" & dfecha1 & "','" & dfecha2 & "','" & dfecha3 & "','" & Trim(txt4.Text) & "','" & Trim(Me.txt_purchar.Text) & "'"

    'sSql = VBsprintf(sSql, "U", sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, DTPicker1.value, DTPicker2.value, DTPicker3.value, txt4.Text)
    ExecuteCommandSQL cCONNECT, sSQl

    Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
    Unload Me

    Exit Sub

errx:
    Err.Raise Err.Number, Err.source, Err.Description
End Sub

Private Sub txt_purchar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Me.funTemCli.SetFocus
    End If

End Sub

Private Sub txt4_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Me.txt_purchar.SetFocus
    End If

End Sub
