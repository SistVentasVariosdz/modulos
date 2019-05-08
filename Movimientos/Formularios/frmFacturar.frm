VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Begin VB.Form frmFacturar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturar"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   3360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraopciones 
      Height          =   795
      Left            =   30
      TabIndex        =   1
      Top             =   1395
      Width           =   3330
      Begin Mantenimientos.MantFunc MantFunc1 
         Height          =   540
         Left            =   1215
         TabIndex        =   8
         Top             =   165
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   953
         Custom          =   $"frmFacturar.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
      End
   End
   Begin VB.Frame fraDatos 
      Height          =   1395
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   3315
      Begin VB.TextBox txtUsuario_Valorizo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1380
         TabIndex        =   4
         Top             =   915
         Width           =   1740
      End
      Begin VB.TextBox txtNum_Docum 
         Height          =   285
         Left            =   1380
         MaxLength       =   15
         TabIndex        =   3
         Top             =   570
         Width           =   1740
      End
      Begin VB.TextBox txtSer_Docum 
         Height          =   285
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   2
         Top             =   225
         Width           =   1740
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Left            =   255
         TabIndex        =   7
         Top             =   975
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Factura"
         Height          =   195
         Left            =   255
         TabIndex        =   6
         Top             =   660
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         Height          =   195
         Left            =   270
         TabIndex        =   5
         Top             =   345
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmFacturar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public varCOD_ALMACEN As String
Public varNUM_MOVSTK As String

Dim varSer_Docum As String
Dim varNum_Docum As String

Dim StrSql As String

Public Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True

    If Trim(Me.txtSer_Docum.Text) = "" And Trim(Me.txtNum_Docum.Text) = "" Then
        varSer_Docum = "Null"
        varNum_Docum = "Null"
    Else
        If Trim(Me.txtSer_Docum.Text) <> "" And Trim(Me.txtNum_Docum.Text) <> "" Then
            varSer_Docum = "'" & Trim(Me.txtSer_Docum.Text) & "'"
            varNum_Docum = "'" & Trim(Me.txtNum_Docum.Text) & "'"
        Else
            VALIDA_DATOS = False
            MsgBox "Los Datos se encuentran imcompletos. Sirvase verificar", vbInformation, "Mensaje"
            Exit Function
        End If
    End If

End Function

Public Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Con.ConnectionString = cConnect
    Con.Open
    
    Con.BeginTrans
        
        StrSql = "EXEC UP_ASIGNAFACTURAMOVISTK '" & _
                Me.varCOD_ALMACEN & "','" & _
                Me.varNUM_MOVSTK & " '," & _
                varSer_Docum & "," & _
                varNum_Docum & ",'" & _
                vusu & "'"
        
    Con.Execute StrSql

    Con.CommitTrans
'    Dim amensaje As New clsMessages
'    amensaje.Codigo = CodeMsg.kMSG_INF_DATA_SAVE
'    Informa "", amensaje
    
    'LIMPIAR_DATOS
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler err, "Salvar_Datos"
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ADICIONAR"
                        If VALIDA_DATOS = True Then
                            Call SALVAR_DATOS
                            Unload Me
                        End If
        Case "SALIR"
                        Unload Me
    End Select
   
End Sub
