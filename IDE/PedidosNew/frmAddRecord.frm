VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form frmAddRecord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar"
   ClientHeight    =   5940
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Registry"
   Begin VB.Frame fraBanco 
      Caption         =   "Registro de Nuevo Banco"
      Height          =   1845
      Left            =   6045
      TabIndex        =   28
      Top             =   4005
      Width           =   5820
      Begin VB.TextBox txtNom_Banco 
         Height          =   285
         Left            =   1620
         MaxLength       =   30
         TabIndex        =   16
         Top             =   615
         Width           =   4035
      End
      Begin VB.TextBox txtCod_Banco 
         Height          =   285
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   15
         Top             =   255
         Width           =   600
      End
      Begin FunctionsButtons.FunctButt funBanco 
         Height          =   510
         Left            =   1635
         TabIndex        =   17
         Top             =   1185
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmAddRecord.frx":0000
         Orientacion     =   0
         Style           =   1
         Language        =   1
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label labels 
         Caption         =   "Banco"
         Height          =   255
         Index           =   10
         Left            =   105
         TabIndex        =   35
         Top             =   285
         Width           =   1335
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
         Index           =   11
         Left            =   135
         TabIndex        =   29
         Tag             =   "Description"
         Top             =   660
         Width           =   945
      End
   End
   Begin VB.Frame fraTipEmb 
      Caption         =   "Registro de Nuevo Embarque"
      Height          =   1845
      Left            =   6045
      TabIndex        =   26
      Top             =   2010
      Width           =   5820
      Begin VB.TextBox txtDes_Embarque 
         Height          =   285
         Left            =   1620
         MaxLength       =   30
         TabIndex        =   13
         Top             =   645
         Width           =   4035
      End
      Begin VB.TextBox txtCod_Embarque 
         Height          =   285
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   12
         Top             =   285
         Width           =   600
      End
      Begin FunctionsButtons.FunctButt funEmbarque 
         Height          =   510
         Left            =   1650
         TabIndex        =   14
         Top             =   1185
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmAddRecord.frx":0097
         Orientacion     =   0
         Style           =   1
         Language        =   1
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label labels 
         Caption         =   "Tipo de Embarque"
         Height          =   255
         Index           =   8
         Left            =   105
         TabIndex        =   34
         Top             =   315
         Width           =   1335
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
         Index           =   9
         Left            =   135
         TabIndex        =   27
         Tag             =   "Description"
         Top             =   660
         Width           =   945
      End
   End
   Begin VB.Frame fraPagEmb 
      Caption         =   "Registro de Nuevo Pago Embarque"
      Height          =   1845
      Left            =   6015
      TabIndex        =   24
      Top             =   15
      Width           =   5820
      Begin VB.TextBox txtDes_PagEmb 
         Height          =   285
         Left            =   1605
         MaxLength       =   30
         TabIndex        =   10
         Top             =   630
         Width           =   4035
      End
      Begin VB.TextBox txtCod_PagEmb 
         Height          =   285
         Left            =   1605
         MaxLength       =   3
         TabIndex        =   9
         Top             =   240
         Width           =   600
      End
      Begin FunctionsButtons.FunctButt funPagEmb 
         Height          =   510
         Left            =   1635
         TabIndex        =   11
         Top             =   1185
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmAddRecord.frx":012E
         Orientacion     =   0
         Style           =   1
         Language        =   1
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label labels 
         Caption         =   "Pago de  Embarque"
         Height          =   255
         Index           =   7
         Left            =   105
         TabIndex        =   33
         Top             =   255
         Width           =   1440
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
         Left            =   135
         TabIndex        =   25
         Tag             =   "Description"
         Top             =   660
         Width           =   945
      End
   End
   Begin VB.Frame fraTemCli 
      Caption         =   "Registro de Nueva Temporada"
      Height          =   1845
      Left            =   30
      TabIndex        =   22
      Top             =   3975
      Width           =   5820
      Begin VB.TextBox txtNom_TemCli 
         Height          =   285
         Left            =   1590
         MaxLength       =   30
         TabIndex        =   7
         Top             =   660
         Width           =   4035
      End
      Begin VB.TextBox txtCod_TemCli 
         Height          =   285
         Left            =   1590
         MaxLength       =   3
         TabIndex        =   6
         Top             =   300
         Width           =   600
      End
      Begin FunctionsButtons.FunctButt funTemCli 
         Height          =   510
         Left            =   1635
         TabIndex        =   8
         Top             =   1185
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmAddRecord.frx":01C5
         Orientacion     =   0
         Style           =   1
         Language        =   1
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label labels 
         Caption         =   "Temporada"
         Height          =   255
         Index           =   5
         Left            =   105
         TabIndex        =   32
         Top             =   330
         Width           =   1335
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
         Left            =   135
         TabIndex        =   23
         Tag             =   "Description"
         Top             =   660
         Width           =   945
      End
   End
   Begin VB.Frame fraDivCli 
      Caption         =   "Registro de Nueva División"
      Height          =   1845
      Left            =   30
      TabIndex        =   20
      Top             =   1995
      Width           =   5820
      Begin VB.TextBox txtCod_DivCli 
         Height          =   285
         Left            =   1590
         MaxLength       =   3
         TabIndex        =   3
         Top             =   270
         Width           =   615
      End
      Begin VB.TextBox txtNom_DivCli 
         Height          =   285
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   4
         Top             =   660
         Width           =   4035
      End
      Begin FunctionsButtons.FunctButt funDivCli 
         Height          =   510
         Left            =   1635
         TabIndex        =   5
         Top             =   1185
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmAddRecord.frx":025C
         Orientacion     =   0
         Style           =   1
         Language        =   1
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label labels 
         Caption         =   "División del Cliente"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   31
         Top             =   315
         Width           =   1335
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
         Index           =   3
         Left            =   135
         TabIndex        =   21
         Tag             =   "Description"
         Top             =   660
         Width           =   945
      End
   End
   Begin VB.Frame fraDestino 
      Caption         =   "Registro de Nuevo Destino"
      Height          =   1845
      Left            =   15
      TabIndex        =   18
      Top             =   0
      Width           =   5820
      Begin VB.TextBox txtCod_Destino 
         Height          =   285
         Left            =   1590
         MaxLength       =   3
         TabIndex        =   0
         Top             =   270
         Width           =   615
      End
      Begin VB.TextBox txtDes_Destino 
         Height          =   285
         Left            =   1590
         MaxLength       =   30
         TabIndex        =   1
         Top             =   630
         Width           =   4050
      End
      Begin FunctionsButtons.FunctButt funDestino 
         Height          =   510
         Left            =   1635
         TabIndex        =   2
         Top             =   1185
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmAddRecord.frx":02F3
         Orientacion     =   0
         Style           =   1
         Language        =   1
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label labels 
         Caption         =   "Destino"
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   30
         Top             =   300
         Width           =   1200
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
         Left            =   135
         TabIndex        =   19
         Tag             =   "Description"
         Top             =   660
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmAddRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oParent        As Object

Public bOk            As Boolean

Public sCod_Cliente   As String

Public sDato          As String

Public sFrame         As String

Public bEnabledCodigo As Boolean

Public Sub LoadFrame(ByRef fraData As Frame)
    fraData.Top = fraDestino.Top
    fraData.Left = fraDestino.Left
    Me.Height = 2200
    Me.Width = 6000
    sFrame = fraData.Name
End Sub

Public Sub LoadFrame2()

    Select Case UCase(sFrame)

        Case "FRADESTINO"
            fraDestino.Top = fraDestino.Top
            fraDestino.Left = fraDestino.Left
            Me.Height = 2200
            Me.Width = 6000

        Case "FRADIVCLI"
            fraDivCli.Top = fraDestino.Top
            fraDivCli.Left = fraDestino.Left
            Me.Height = 2200
            Me.Width = 6000

        Case "FRATEMCLI"
            fraTemCli.Top = fraDestino.Top
            fraTemCli.Left = fraDestino.Left
            Me.Height = 2200
            Me.Width = 6000

        Case "FRAPAGEMB"
            fraPagEmb.Top = fraDestino.Top
            fraPagEmb.Left = fraDestino.Left
            Me.Height = 2200
            Me.Width = 6000

        Case "FRATIPEMB"
            fraTipEmb.Top = fraDestino.Top
            fraTipEmb.Left = fraDestino.Left
            Me.Height = 2200
            Me.Width = 6000

        Case "FRABANCO"
            fraBanco.Top = fraDestino.Top
            fraBanco.Left = fraDestino.Left
            Me.Height = 2200
            Me.Width = 6000
    End Select

End Sub

Private Sub Form_Activate()

    Select Case UCase(sFrame)

        Case "FRADESTINO"

            If bEnabledCodigo Or RTrim(txtCod_Destino) = "" Then
                If txtCod_Destino.Enabled Then
                    txtCod_Destino.SetFocus
                End If

            Else

                If txtDes_Destino.Enabled Then
                    txtDes_Destino.SetFocus
                End If
            End If

        Case "FRADIVCLI"

            If bEnabledCodigo Or RTrim(txtCod_DivCli) = "" Then
                If txtCod_DivCli.Enabled Then
                    txtCod_DivCli.SetFocus
                End If

            Else

                If txtNom_DivCli.Enabled Then
                    txtNom_DivCli.SetFocus
                End If
            End If

        Case "FRATEMCLI"

            If bEnabledCodigo Or RTrim(txtCod_TemCli) = "" Then
                If txtCod_TemCli.Enabled Then
                    txtCod_TemCli.SetFocus
                End If

            Else

                If txtNom_TemCli.Enabled Then
                    txtNom_TemCli.SetFocus
                End If
            End If

        Case "FRAPAGEMB"

            If bEnabledCodigo Or RTrim(txtCod_PagEmb) = "" Then
                If txtCod_PagEmb.Enabled Then
                    txtCod_PagEmb.SetFocus
                End If

            Else

                If txtDes_PagEmb.Enabled Then
                    txtDes_PagEmb.SetFocus
                End If
            End If

        Case "FRATIPEMB"

            If bEnabledCodigo Or RTrim(txtCod_Embarque) = "" Then
                If txtCod_Embarque.Enabled Then
                    txtCod_Embarque.SetFocus
                End If

            Else

                If txtDes_Embarque.Enabled Then
                    txtDes_Embarque.SetFocus
                End If
            End If

        Case "FRABANCO"

            If bEnabledCodigo Or RTrim(txtCod_Banco) = "" Then
                If txtCod_Banco.Enabled Then
                    txtCod_Banco.SetFocus
                End If

            Else

                If txtNom_Banco.Enabled Then
                    txtNom_Banco.SetFocus
                End If
            End If

    End Select

End Sub

Private Sub funDestino_ActionClick(ByVal Index As Integer, _
                                   ByVal ActionType As Integer, _
                                   ByVal ActionName As String)

    On Error GoTo errx

    Dim obj As clsTG_LotColTal

    Select Case ActionName

        Case "ACEPTAR"

            If RTrim(txtCod_Destino.Text) = "" Then
                Mensaje kMESSAGE_ERR_NOTEMPTY

                Exit Sub

            End If

            If RTrim(txtDes_Destino.Text) = "" Then
                Mensaje kMESSAGE_ERR_NOTEMPTY

                Exit Sub

            End If

            Set obj = New clsTG_LotColTal
            obj.ConexionString = cCONNECT
            obj.AddDestino txtCod_Destino.Text, txtDes_Destino.Text
            Set obj = Nothing
            bOk = True
            sDato = txtCod_Destino.Text
            Unload Me

        Case "CANCELAR"
            Unload Me
    End Select

    Exit Sub

errx:

    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
End Sub

Private Sub funDivCli_ActionClick(ByVal Index As Integer, _
                                  ByVal ActionType As Integer, _
                                  ByVal ActionName As String)

    On Error GoTo errx

    Dim obj As clsTG_LotColTal

    Select Case ActionName

        Case "ACEPTAR"

            If RTrim(txtCod_DivCli.Text) = "" Then
                Mensaje kMESSAGE_ERR_NOTEMPTY

                Exit Sub

            End If

            If RTrim(txtNom_DivCli.Text) = "" Then
                Mensaje kMESSAGE_ERR_NOTEMPTY

                Exit Sub

            End If

            Set obj = New clsTG_LotColTal
            obj.ConexionString = cCONNECT
            obj.AddDivCli sCod_Cliente, txtCod_DivCli.Text, txtNom_DivCli.Text
            Set obj = Nothing
            bOk = True
            sDato = txtCod_DivCli.Text
            Unload Me

        Case "CANCELAR"
            Unload Me
    End Select

    Exit Sub

errx:

    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
End Sub

Private Sub funTemCli_ActionClick(ByVal Index As Integer, _
                                  ByVal ActionType As Integer, _
                                  ByVal ActionName As String)

    On Error GoTo errx

    Dim obj As clsTG_LotColTal

    Select Case ActionName

        Case "ACEPTAR"

            If RTrim(txtCod_TemCli.Text) = "" Then
                Mensaje kMESSAGE_ERR_NOTEMPTY

                Exit Sub

            End If

            If RTrim(txtNom_TemCli.Text) = "" Then
                Mensaje kMESSAGE_ERR_NOTEMPTY

                Exit Sub

            End If

            Set obj = New clsTG_LotColTal
            obj.ConexionString = cCONNECT
            obj.AddTemCli sCod_Cliente, txtCod_TemCli.Text, txtNom_TemCli.Text
            Set obj = Nothing
            bOk = True
            sDato = txtCod_TemCli.Text
            Unload Me

        Case "CANCELAR"
            Unload Me
    End Select

    Exit Sub

errx:

    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
End Sub

Private Sub funPagEmb_ActionClick(ByVal Index As Integer, _
                                  ByVal ActionType As Integer, _
                                  ByVal ActionName As String)

    On Error GoTo errx

    Dim obj As clsTG_LotColTal

    Select Case ActionName

        Case "ACEPTAR"

            If RTrim(txtCod_PagEmb.Text) = "" Then
                Mensaje kMESSAGE_ERR_NOTEMPTY

                Exit Sub

            End If

            If RTrim(txtDes_PagEmb.Text) = "" Then
                Mensaje kMESSAGE_ERR_NOTEMPTY

                Exit Sub

            End If

            Set obj = New clsTG_LotColTal
            obj.ConexionString = cCONNECT
            obj.AddPagEmb txtCod_PagEmb.Text, txtDes_PagEmb.Text
            Set obj = Nothing
            bOk = True
            sDato = txtCod_PagEmb.Text
            Unload Me

        Case "CANCELAR"
            Unload Me
    End Select

    Exit Sub

errx:

    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
End Sub

Private Sub funEmbarque_ActionClick(ByVal Index As Integer, _
                                    ByVal ActionType As Integer, _
                                    ByVal ActionName As String)

    On Error GoTo errx

    Dim obj As clsTG_LotColTal

    Select Case ActionName

        Case "ACEPTAR"

            If RTrim(txtCod_Embarque.Text) = "" Then
                Mensaje kMESSAGE_ERR_NOTEMPTY

                Exit Sub

            End If

            If RTrim(txtDes_Embarque.Text) = "" Then
                Mensaje kMESSAGE_ERR_NOTEMPTY

                Exit Sub

            End If

            Set obj = New clsTG_LotColTal
            obj.ConexionString = cCONNECT
            obj.AddTipEmb txtCod_Embarque.Text, txtDes_Embarque.Text
            Set obj = Nothing
            bOk = True
            sDato = txtCod_Embarque.Text
            Unload Me

        Case "CANCELAR"
            Unload Me
    End Select

    Exit Sub

errx:

    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
End Sub

Private Sub funBanco_ActionClick(ByVal Index As Integer, _
                                 ByVal ActionType As Integer, _
                                 ByVal ActionName As String)

    On Error GoTo errx

    Dim obj As clsTG_LotColTal

    Select Case ActionName

        Case "ACEPTAR"

            If RTrim(txtCod_Banco.Text) = "" Then
                Mensaje kMESSAGE_ERR_NOTEMPTY

                Exit Sub

            End If

            If RTrim(txtNom_Banco.Text) = "" Then
                Mensaje kMESSAGE_ERR_NOTEMPTY

                Exit Sub

            End If

            Set obj = New clsTG_LotColTal
            obj.ConexionString = cCONNECT
            obj.AddBanco txtCod_Banco.Text, txtNom_Banco.Text
            Set obj = Nothing
            bOk = True
            sDato = txtCod_Banco.Text
            Unload Me

        Case "CANCELAR"
            Unload Me
    End Select

    Exit Sub

errx:

    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
End Sub

