VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Actualiza_FechaPO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualiza Fecha Llegada PO"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5880
      TabIndex        =   4
      Top             =   930
      Width           =   1245
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4620
      TabIndex        =   3
      Top             =   930
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   7095
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2370
         TabIndex        =   1
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   51970051
         CurrentDate     =   40263.5095949074
      End
      Begin VB.Label Label2 
         Caption         =   "(Formato Hora 00 -23)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4710
         TabIndex        =   5
         Top             =   420
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Llegada PO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   2
         Top             =   420
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Frm_Actualiza_FechaPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sCod_Cliente As String
Public sCod_PurOrd As String
Public oParent As Object

Private Sub cmdAceptar_Click()
Dim SSQL As String
If MsgBox("¿Realmente Desea Actualizar la Fecha?", vbYesNo, "Mensaje del Sistema") = vbYes Then
    If FixNulos(Me.DTPicker1.value, vbstring) = "" Then
        sFec_Hora_LLeg = "NULL"
        SSQL = "EXEC TG_PURORD_ACTUALIZA_FEC_HORA_LLEGADA_PO '$','$',$"
        SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_PurOrd, sFec_Hora_LLeg)
        ExecuteCommandSQL cCONNECT, SSQL
        Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
        oParent.BUSCAR
        oParent.BuscarEStilos
        Unload Me
    Else
        sFec_Hora_LLeg = Format(Me.DTPicker1.value, "dd/MM/yyyy HH:MM")
                     
        SSQL = "EXEC TG_PURORD_ACTUALIZA_FEC_HORA_LLEGADA_PO '$','$','$'"
        SSQL = VBsprintf(SSQL, sCod_Cliente, sCod_PurOrd, sFec_Hora_LLeg)
        ExecuteCommandSQL cCONNECT, SSQL
        Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
        oParent.BUSCAR
        oParent.BuscarEStilos
        Unload Me
    End If
Else
    Unload Me
End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

