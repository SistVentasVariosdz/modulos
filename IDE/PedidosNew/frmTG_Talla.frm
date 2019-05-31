VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form frmTG_Talla 
   Caption         =   "Creación de Talla"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3495
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Add Size"
   Begin VB.TextBox txtCod_Talla 
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
      Left            =   975
      MaxLength       =   10
      TabIndex        =   0
      Top             =   105
      Width           =   2310
   End
   Begin FunctionsButtons.FunctButt acbForm 
      Height          =   510
      Left            =   465
      TabIndex        =   2
      Top             =   720
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmTG_Talla.frx":0000
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
      Caption         =   "Talla :"
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
      TabIndex        =   1
      Tag             =   "Size"
      Top             =   135
      Width           =   705
   End
End
Attribute VB_Name = "frmTG_Talla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bOk         As Boolean

Public sCod_Talla  As String

Public sCod_GruTal As String

Private Sub acbForm_ActionClick(ByVal Index As Integer, _
                                ByVal ActionType As Integer, _
                                ByVal ActionName As String)

    On Error GoTo errores

    Dim vbuff

    Dim obj       As clsTG_Talla

    Dim sTipColor As String
    
    Select Case ActionName

        Case "ACEPTAR"

            If txtCod_Talla.Text = "" Then
                If txtCod_Talla.Enabled Then
                    txtCod_Talla.SetFocus
                End If

                Exit Sub

            End If
        
            Set obj = New clsTG_Talla
            obj.ConexionString = cCONNECT
            obj.Add Me.txtCod_Talla, sCod_GruTal
            Set obj = Nothing
            sCod_Talla = Me.txtCod_Talla.Text
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

