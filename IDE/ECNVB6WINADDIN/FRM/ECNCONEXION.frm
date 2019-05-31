VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "*\A..\..\ECNVB6WINCTRL\ECNVB6WINCTRL.vbp"
Begin VB.Form frm003_ECNCONEXION 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Datos de Conexión"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ECNCONEXION.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrWaitPbr 
      Interval        =   1000
      Left            =   900
      Top             =   3990
   End
   Begin VB.PictureBox pctCabecera 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFE8DD&
      DrawWidth       =   3
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   0
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   365
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   5505
      Begin ECNVB6WINCTRL.ucLabel lblSistema 
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   540
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   873
         Caption         =   "Sistemas de Información"
         ForeColor       =   0
         BackColor       =   15722717
         ShadowColor     =   6710886
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ECNVB6WINCTRL.ucLabel lblEmpresa 
         Height          =   375
         Left            =   90
         TabIndex        =   19
         Top             =   150
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   661
         Caption         =   "ECNSoft"
         ForeColor       =   13421772
         BackColor       =   15722717
         ShadowColor     =   3342336
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer tmrProgressBar 
      Interval        =   50
      Left            =   450
      Top             =   3990
   End
   Begin VB.Timer tmrConexion 
      Interval        =   1000
      Left            =   0
      Top             =   3990
   End
   Begin ECNVB6WINCTRL.ucButton_02 btnConectar 
      Height          =   375
      Left            =   2790
      TabIndex        =   13
      Top             =   3390
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      Icon            =   "ECNCONEXION.frx":058A
      Style           =   5
      Caption         =   "    &Conectar"
      iNonThemeStyle  =   0
      Object.ToolTipText     =   ""
      ToolTipTitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin VB.PictureBox pctMsg 
      Height          =   345
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   2685
      TabIndex        =   11
      Top             =   3390
      Width           =   2745
      Begin ECNVB6WINCTRL.ucLabel lblProceso 
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   60
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   556
         Caption         =   "Falta Definir Conexion..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ECNVB6WINCTRL.ucImage icnConnect 
         Height          =   240
         Left            =   2400
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   423
         bvData          =   "ECNCONEXION.frx":0B24
         bData           =   -1  'True
         Filename        =   "connect16_d.ico"
         eScale          =   0
         lContrast       =   0
         lBrightness     =   0
         lAlpha          =   100
         bGrayScale      =   0   'False
         lAngle          =   0
         bFlipH          =   -1  'True
         bFlipV          =   0   'False
      End
   End
   Begin ECNVB6WINCTRL.ucButton_02 btnSalir 
      Height          =   375
      Left            =   4170
      TabIndex        =   14
      Top             =   3390
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Icon            =   "ECNCONEXION.frx":153A
      Style           =   5
      Caption         =   "    &Cancelar"
      iNonThemeStyle  =   0
      Object.ToolTipText     =   ""
      ToolTipTitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin ECNVB6WINCTRL.ucProgressBar prbConectar 
      Height          =   105
      Left            =   0
      TabIndex        =   21
      Top             =   1380
      Visible         =   0   'False
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   185
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   16750899
      Scrolling       =   2
   End
   Begin VB.Frame famDatos 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   0
      TabIndex        =   5
      Top             =   1290
      Width           =   5505
      Begin ECNVB6WINCTRL.ucButton_02 btnRedWin 
         Height          =   315
         Left            =   5010
         TabIndex        =   17
         ToolTipText     =   "Muestra una ventana que realiza la busqueda de equipos cercanos dentro de la red de Windows apoyandose del algoritmo de Windows"
         Top             =   270
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         Icon            =   "ECNCONEXION.frx":1AD4
         iNonThemeStyle  =   0
         Object.ToolTipText     =   "Muestra una ventana que realiza la busqueda de equipos cercanos dentro de la red de Windows apoyandose del algoritmo de Windows"
         ToolTipTitle    =   "Búsqueda de Servidores"
         ToolTipIcon     =   1
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin VB.TextBox txtServidor 
         BackColor       =   &H00F8F8F8&
         Height          =   315
         Left            =   1710
         TabIndex        =   0
         Top             =   270
         Width           =   2535
      End
      Begin MSDataListLib.DataCombo dcBD 
         Height          =   315
         Left            =   1710
         TabIndex        =   3
         Top             =   1260
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtUsuario 
         BackColor       =   &H00F8F8F8&
         Height          =   315
         Left            =   1710
         TabIndex        =   1
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtClave 
         BackColor       =   &H00F8F8F8&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "l"
         TabIndex        =   2
         Top             =   915
         Width           =   3615
      End
      Begin MSDataListLib.DataCombo dcBDI 
         Height          =   315
         Left            =   1710
         TabIndex        =   4
         Top             =   1590
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ECNVB6WINCTRL.ucButton_02 btnLocal 
         Height          =   315
         Left            =   4260
         TabIndex        =   15
         ToolTipText     =   "Definir como servidor SQL el equipo local"
         Top             =   270
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         Icon            =   "ECNCONEXION.frx":24E6
         iNonThemeStyle  =   0
         Object.ToolTipText     =   "Definir como servidor SQL el equipo local"
         ToolTipTitle    =   "Búsqueda de Servidores"
         ToolTipIcon     =   1
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin ECNVB6WINCTRL.ucButton_02 btnRed 
         Height          =   315
         Left            =   4650
         TabIndex        =   16
         ToolTipText     =   "Muestra una ventana que realiza la bsuqueda de equipos cercanos dentro de la red de Windows"
         Top             =   270
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Icon            =   "ECNCONEXION.frx":2A80
         iNonThemeStyle  =   0
         Object.ToolTipText     =   "Muestra una ventana que realiza la bsuqueda de equipos cercanos dentro de la red de Windows"
         ToolTipTitle    =   "Búsqueda de Servidores"
         ToolTipIcon     =   1
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   180
         Picture         =   "ECNCONEXION.frx":301A
         Stretch         =   -1  'True
         ToolTipText     =   "Ruta en el Servidor de BD, donde se guardará la nueva Base de Datos..."
         Top             =   330
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   180
         Picture         =   "ECNCONEXION.frx":3404
         Stretch         =   -1  'True
         ToolTipText     =   "Ruta en el Servidor de BD, donde se guardará la nueva Base de Datos..."
         Top             =   660
         Width           =   240
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   180
         Picture         =   "ECNCONEXION.frx":363C
         Stretch         =   -1  'True
         ToolTipText     =   "Ruta en el Servidor de BD, donde se guardará la nueva Base de Datos..."
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   180
         Picture         =   "ECNCONEXION.frx":3A39
         Stretch         =   -1  'True
         Top             =   1350
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Servidor"
         Height          =   195
         Left            =   510
         TabIndex        =   10
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Left            =   510
         TabIndex        =   9
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña"
         Height          =   195
         Left            =   510
         TabIndex        =   8
         Top             =   1035
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Base de Datos"
         Height          =   195
         Left            =   510
         TabIndex        =   7
         Top             =   1380
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "B.Datos Imágen"
         Height          =   195
         Left            =   510
         TabIndex        =   6
         Top             =   1710
         Width           =   1155
      End
      Begin ECNVB6WINCTRL.ucImageList imlConexion 
         Left            =   4830
         Top             =   1950
         _ExtentX        =   1005
         _ExtentY        =   1005
         bvData          =   "ECNCONEXION.frx":3E22
         bData           =   -1  'True
      End
   End
End
Attribute VB_Name = "frm003_ECNCONEXION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'************************************************************************************************************************************************************************************************************************************************
'VARIABLES PUBLICAS DEL PROGRAMA
'************************************************************************************************************************************************************************************************************************************************

'************************************************************************************************************************************************************************************************************************************************
'Api para mantener siempre visible el Formulario
'************************************************************************************************************************************************************************************************************************************************
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wPU_003_SW_Conexions As Long) As Long

'--+--------------------------------------------+--
'Constants for SetWindowPos
'--+--------------------------------------------+--
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
'************************************************************************************************************************************************************************************************************************************************
' VARIABLES PARA EL MOVIMIENTO DEL FORMULARIO
'************************************************************************************************************************************************************************************************************************************************
Private DifferenceX As Single
Private DifferenceY As Single

Private StartMoving As Boolean
Private WindUp As Boolean
Private KeyCount As Integer
'
'Private blSW_EfectoCabecera_End As Boolean
'Private blSW_ProgressBarCon_End As Boolean

Private Sub Form_Load()
    GO_003_PU_SW_LOAD = True
'    blSW_EfectoCabecera_End = False
'    blSW_ProgressBarCon_End = False
    With PU_003_ECNLIB01_FUNSUB
        Call .DeshabilitarBotonXdeForm(Me)
    End With
    StartMoving = False
    WindUp = False
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    GO_003_PU_SW_LOAD = False
End Sub

Private Sub Form_Activate()
    If txtServidor.Visible = True Then txtServidor.SetFocus
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    DoMouseDown Button, Shift, X, y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    DoMouseMove Button, Shift, X, y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    DoMouseUp Button, Shift, X, y
End Sub

Private Sub btnLocal_Click()
    txtServidor.Text = GO_ECNLIB01_FUNSUB.ComputerName
    txtServidor.SetFocus
End Sub

Private Sub btnRed_Click()
    Dim oECNNET As New cls004_ECNNET
    
    With oECNNET
    End With
    With oECNNET
        .FORM_PARENT = Me
        With .GetForm
            .Top = Me.Top - 1000
            .Left = Me.Left + Me.Width + 200
        End With
        .ShowPrompt
        txtServidor.Text = .COMPUTER_SEL
    End With
    Set oECNNET = Nothing
    If Len(Trim(txtServidor.Text)) > 0 Then
        txtUsuario.SetFocus
    Else
        txtServidor.SetFocus
    End If
End Sub

Private Sub btnRedWin_Click()
    txtServidor.Text = PU_003_ECNLIB01_FUNSUB.BuscarEquiposDeRed(Me)
    txtServidor.SetFocus
End Sub

Private Sub icnConnect_MouseEnter()
    icnConnect.Angle = 10
End Sub

Private Sub icnConnect_MouseExit()
    icnConnect.Angle = 0
End Sub

Private Sub tmrConexion_Timer()
    If Me.Visible = True Then
        Call PU_003_ECNLIB04_EFFECTS.LetrasEnCaidaDelluvias("Desarrollado por Edgar I. Cárdenas Navarro", _
                                                            pctCabecera, _
                                                            0)
        tmrConexion.Interval = 0
        tmrConexion.Enabled = False
    End If
End Sub

Private Sub tmrProgressBar_Timer()
    If prbConectar.Value >= prbConectar.Max Then
        'tmrProgressBar.Enabled = False
        prbConectar.Value = 0
        Exit Sub
    End If
    prbConectar.Value = prbConectar.Value + 1
    DoEvents
End Sub

'Private Sub tmrWaitPbr_Timer()
'    Static iConSegundos As Integer
'
'    iConSegundos = iConSegundos + 1
'    if iConSegundos
'End Sub

Private Sub txtServidor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtServidor) = 0 Then
            txtServidor.SetFocus
        Else
            txtUsuario.SetFocus
        End If
    End If
End Sub

Private Sub txtServidor_GotFocus()
    Call PU_003_ECNLIB03_WINEVE.Ctrl_GotFocus(txtServidor)
End Sub

Private Sub txtUsuario_GotFocus()
    Call PU_003_ECNLIB03_WINEVE.Ctrl_GotFocus(txtUsuario)
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtClave.SetFocus
End Sub

Private Sub txtClave_GotFocus()
    Call PU_003_ECNLIB03_WINEVE.Ctrl_GotFocus(txtClave)
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then btnConectar.SetFocus
End Sub

Private Sub btnConectar_Click()
    Call ActivarTmrProgressBar(True)
    If btnConectar.Caption = "&Aceptar" Then
        Select Case GO_003_MODO_WIN
            Case WINC_Main
                Call PU_003_CrearECNOVL(PU_003_SERVER, _
                                        PU_003_USUARIO, _
                                        PU_003_CLAVE, _
                                        PU_003_BD, _
                                        PU_003_BDI)
            Case WinC_Prompt
                If GO_003_FORM_PARENT Is Nothing Then
                    Call ActivarTmrProgressBar(False)
                    MsgBox "No se ha definido el formulario que recibirá los datos del modo PROMPT", _
                           vbCritical, _
                           Me.Caption
                    Exit Sub
                End If
                With GO_003_FORM_PARENT
                    .PU_SQLSERV = PU_003_SERVER
                    .PU_SQLUSER = PU_003_USUARIO
                    .PU_SQLUPWD = PU_003_CLAVE
                    .PU_NBDMAIN = PU_003_BD
                    .PU_NBDIMAG = PU_003_BDI
                    .PU_SQLRUCN = PU_003_CONNECT
                    
                    .PU_WCONCAN = False
                End With
        End Select
        GO_003_ENU_OPC_WIN_RESULT = WD_ACCEPT
        GoTo Fin
    Else
        lblProceso.Caption = "Verificando..."
        Call ADO
    End If
    If PU_003_SW_Conexion = True Then Call BD
    Call ActivarTmrProgressBar(False)
    Exit Sub
Fin:
    Unload Me
End Sub

Public Sub dcBD_Click(Area As Integer)
    Call PU_003_Click_End_dcBD(Area)
End Sub

Private Sub dcBD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcBD_Click(2)
        dcBDI.SetFocus
    End If
End Sub

Private Sub dcBDI_Click(Area As Integer)
    If Area <> 2 Then Exit Sub
    PU_003_BDI = dcBDI.Text
End Sub

Private Sub dcBDI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then btnConectar.SetFocus
End Sub

Private Sub btnSalir_Click()
    Select Case GO_003_MODO_WIN
        Case WINC_Main
            Unload Me
        Case WinC_Prompt
            If Not (GO_003_FORM_PARENT Is Nothing) Then GO_003_FORM_PARENT.PU_WCONCAN = True
            Unload Me
    End Select
End Sub

'************************************************************************************************************************************************************************************************************************************************
' PROCEDIMIENTOS Y FUNCIONES DE USUARIO
'************************************************************************************************************************************************************************************************************************************************
Public Sub ADO()
    Call PU_003_ADO
End Sub

Public Sub BD()
   Call PU_003_GetBD
End Sub

Private Sub ActivarTmrProgressBar(ByVal blSW As Boolean)
    Select Case blSW
        Case True
            famDatos.Top = 1410
            pctMsg.Top = 3540
            Me.Height = 4350
        Case False
            famDatos.Top = 1290
            pctMsg.Top = 3390
            Me.Height = 4200
    End Select
    With pctMsg
        btnConectar.Top = .Top
        btnSalir.Top = .Top
    End With
    With prbConectar
        .Value = 0
        .Visible = blSW
    End With
    Me.Refresh
    Call PU_003_ECNLIB01_FUNSUB.Esperar(1)
    'tmrProgressBar.Enabled = blSW
End Sub

Private Sub DoMouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        StartMoving = True
        WindUp = False
        DifferenceX = X
        DifferenceY = y
    End If
End Sub

Private Sub DoMouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        If StartMoving Then
            Me.Left = Me.Left + (X - DifferenceX)
            Me.Top = Me.Top + (y - DifferenceY)
            DoEvents
        End If
    End If
End Sub

Private Sub DoMouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If StartMoving And Button = 1 Then
        WindUp = False
        StartMoving = False
        KeyCount = 0
    End If
End Sub

