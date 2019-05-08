VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmDetalleAvios 
   Caption         =   "Detalle de Movimiento"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtobservacion 
         Height          =   320
         Left            =   1920
         TabIndex        =   7
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox txtcantidad 
         Height          =   320
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPFecha 
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   60227585
         CurrentDate     =   37460
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Observación :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   470
         TabIndex        =   4
         Top             =   1300
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   3
         Top             =   840
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   2
         Top             =   390
         Width           =   945
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2310
      TabIndex        =   0
      Top             =   1920
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmDetalleAvios.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmDetalleAvios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Fabrica As String, sCod_OrdPro As String, SCOD_PROVEEDOR As String, sCOD_ITEM As String, sCOD_COMB As String
Public scod_color As String, sCOD_TALLA As String, sCod_Descuento As String, sEstilo_Cliente As String
Public sCantidad_Enviada As Double
Public sql As String
Dim strSQL As String

Private Sub DTPFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtobservacion.SetFocus
    End If
End Sub



Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "MODIFICAR"
    Call guardar
Case "SALIR"
    Unload Me
End Select
End Sub

Sub guardar()

'On Error GoTo hand

If (txtcantidad.Text) < sCantidad_Enviada Then
 sql = "lg_up_man__Es_OrdProReq_Items_Devueltos_Proveedores  '" & Trim(sCod_Fabrica) & "','" & Trim(sCod_OrdPro) & "','" & Trim(SCOD_PROVEEDOR) & "','" & Trim(sCOD_ITEM) & "','" & Trim(sCOD_COMB) & "','" & Trim(scod_color) & "','" & Trim(sCOD_TALLA) & "','" & Trim(sCod_Descuento) & "','" & Trim(sEstilo_Cliente) & "','" & Trim(txtcantidad.Text) & "','" & DTPFecha & "','" & Trim(txtobservacion.Text) & "'"

 Call ExecuteSQL(cConnect, sql)
 MsgBox "Se grabo con éxito"
 FrmShowAviosxServicioConfec.CARGA_GRID
 Unload Me
Exit Sub

Else

MsgBox "La Cantidad debe ser menor a la Cantidad Enviada"
End If
'hand:
'    SALVAR_CABECERA = False
'    ErrorHandler err, "Grabar Conceptos"
End Sub



Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DTPFecha.SetFocus
    End If
End Sub

Private Sub txtobservacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FunctButt1.SetFocus
    End If
End Sub
