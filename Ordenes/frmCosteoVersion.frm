VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmVersionCosteo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio Costeo-Version"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1785
      TabIndex        =   17
      Top             =   2505
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "0~0~ACEPTAR~True~True~&Aceptar~0~0~1~~0~False~False~&Aceptar~~1~0~CANCELAR~True~True~&Cancelar~0~0~2~~0~False~False~&Cancelar~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   2340
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   6150
      Begin VB.TextBox txtDes_Version 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2145
         TabIndex        =   16
         Top             =   1890
         Width           =   3315
      End
      Begin VB.TextBox txtCod_Version 
         Height          =   315
         Left            =   1335
         TabIndex        =   15
         Top             =   1890
         Width           =   795
      End
      Begin VB.TextBox txtDes_EstPro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2145
         TabIndex        =   13
         Top             =   1530
         Width           =   3315
      End
      Begin VB.TextBox txtCod_EstPro 
         Height          =   315
         Left            =   1335
         TabIndex        =   12
         Top             =   1530
         Width           =   795
      End
      Begin VB.TextBox txtCotizacion 
         Height          =   315
         Left            =   1335
         TabIndex        =   10
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label Label7 
         Caption         =   "Version :"
         Height          =   195
         Left            =   195
         TabIndex        =   14
         Top             =   1980
         Width           =   900
      End
      Begin VB.Label Label6 
         Caption         =   "Estilo Propio :"
         Height          =   225
         Left            =   165
         TabIndex        =   11
         Top             =   1635
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Cotizacion :"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   1260
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6045
         Y1              =   1065
         Y2              =   1065
      End
      Begin VB.Label lblEstPro 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   4005
         TabIndex        =   8
         Top             =   675
         Width           =   1920
      End
      Begin VB.Label lblEstCli 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   1050
         TabIndex        =   7
         Top             =   675
         Width           =   1920
      End
      Begin VB.Label lblOP 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   4005
         TabIndex        =   6
         Top             =   225
         Width           =   1920
      End
      Begin VB.Label lblPO 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   1065
         TabIndex        =   5
         Top             =   270
         Width           =   1920
      End
      Begin VB.Label Label4 
         Caption         =   "O.P.:"
         Height          =   225
         Left            =   3195
         TabIndex        =   4
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "Est.Propio :"
         Height          =   240
         Left            =   3180
         TabIndex        =   3
         Top             =   675
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "Est.Cliente :"
         Height          =   225
         Left            =   120
         TabIndex        =   2
         Top             =   690
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "P.O.:"
         Height          =   225
         Left            =   135
         TabIndex        =   1
         Top             =   255
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmVersionCosteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_cliente As String
Public sCod_PurOrd As String
Public sCod_LotPurOrd As String
Public sCod_EstCli As String
Public sCod_EstPro As String

Dim strSQL As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            If Trim(txtCotizacion.Text) = "" Then
                MsgBox "Ingrese el numero de Cotizacion", vbInformation, "Aviso"
                txtCotizacion.SetFocus
                Exit Sub
            End If
        
            If Trim(txtCod_EstPro.Text) = "" Then
                MsgBox "Seleccione el Estilo Propio", vbInformation, "Aviso"
                txtCod_EstPro.SetFocus
                Exit Sub
            End If
        
            If Trim(txtCod_Version.Text) = "" Then
                MsgBox "Seleccione la Version", vbInformation, "Aviso"
                txtCod_EstPro.SetFocus
                Exit Sub
            End If
            SALVAR_DATOS
            
        Case "CANCELAR"
            Unload Me
    End Select
End Sub

Private Sub txtCod_EstPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim oTipo As New frmBusqGrande
        Set oTipo.oParent = Me
        oTipo.SQuery = "exec SM_CONSULTA_ESTILOS_VERSION_POR_COTIZACION " & txtCotizacion.Text
        oTipo.CARGAR_DATOS
        oTipo.Show 1
        If Trim(txtCod_EstPro.Text) <> "" And Trim(txtCod_Version) <> "" Then Me.FunctButt1.SetFocus
    End If
End Sub

Private Sub txtCotizacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtCotizacion.Text = "" Then
            Load frmCotizacionesAlternativas
            frmCotizacionesAlternativas.sCod_cliente = sCod_cliente
            frmCotizacionesAlternativas.sCod_EstCli = sCod_EstCli
            frmCotizacionesAlternativas.Buscar
            frmCotizacionesAlternativas.Show vbModal
            If frmCotizacionesAlternativas.bOk Then
                txtCotizacion = frmCotizacionesAlternativas.vNum_cotizacion
                txtCod_EstPro = frmCotizacionesAlternativas.vCod_Estpro_Cotizacion
                txtDes_EstPro = DevuelveCampo("SELECT DES_ESTPRO FROM ES_ESTPRO WHERE COD_ESTPRO = '" & txtCod_EstPro & "'", cCONNECT)
                txtCod_Version = frmCotizacionesAlternativas.vCod_Version_Cotizacion
                txtDes_Version = DevuelveCampo("SELECT DES_VERSION FROM ES_ESTPROVER WHERE COD_ESTPRO = '" & txtCod_EstPro & "' AND COD_VERSION = '" & txtCod_Version & "'", cCONNECT)
                FunctButt1.SetFocus
            End If
            Set frmCotizacionesAlternativas = Nothing
        Else
            If ExisteCampo("num_solicitud_cons", "tg_cotizacion", txtCotizacion.Text, cCONNECT) Then
                txtCod_EstPro.SetFocus
            Else
                MsgBox "El Estilo no existe", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
    Else
        Call SoloNumeros(txtCotizacion, KeyAscii, False)
    End If
End Sub

Sub SALVAR_DATOS()
On Error GoTo errores

strSQL = "EXEC UP_MAN_TG_LOTESTPRO_VERSIONES_COSTEO '" & sCod_cliente & "','" & _
                                                sCod_PurOrd & "','" & _
                                                sCod_LotPurOrd & "','" & _
                                                sCod_EstCli & "','" & _
                                                sCod_EstPro & "'," & _
                                                txtCotizacion.Text & ",'" & _
                                                txtCod_EstPro.Text & "','" & _
                                                txtCod_Version.Text & "'"
                                                
Call ExecuteCommandSQL(cCONNECT, strSQL)

Unload Me

Exit Sub
errores:
    ErrorHandler Err, "SALVAR_DATOS"
End Sub


