VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSeleccionCotizacionSgt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingrese datos Adicionales"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   3450
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton CmdFin 
      Caption         =   "&Finalizar"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   2280
      Width           =   1515
   End
   Begin VB.TextBox txtObs 
      Height          =   465
      Left            =   840
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1650
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingrese los siguientes datos:"
      Height          =   1430
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   6375
      Begin MSComCtl2.DTPicker FProtos 
         Height          =   315
         Left            =   4440
         TabIndex        =   5
         Top             =   780
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   23920641
         CurrentDate     =   37453
      End
      Begin MSComCtl2.DTPicker FEntrega 
         Height          =   315
         Left            =   2310
         TabIndex        =   4
         Top             =   780
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   23920641
         CurrentDate     =   37376
      End
      Begin MSComCtl2.DTPicker FSolicitud 
         Height          =   345
         Left            =   270
         TabIndex        =   1
         Top             =   750
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         _Version        =   393216
         Format          =   23920641
         CurrentDate     =   37376
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Entrega Protos"
         Height          =   255
         Left            =   4440
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de entrega:"
         Height          =   255
         Left            =   2310
         TabIndex        =   3
         Top             =   510
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de solicitud:"
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1485
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Obs:"
      Height          =   285
      Left            =   270
      TabIndex        =   9
      Top             =   1680
      Width           =   435
   End
End
Attribute VB_Name = "frmSeleccionCotizacionSgt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TipoFab As Integer

Private Sub cmdCancelar_Click()
Dim nMessage As Integer
    nMessage = MsgBox("Esta seguro de cancelar el proceso", vbYesNo, "Proceso de Cotización")
    Select Case nMessage
        Case vbYes
            frmSeleccionCotizacion.varEst_Avance = False
            Unload Me
        Case vbNo
            Exit Sub
    End Select
End Sub

Private Sub CmdFin_Click()
    If ValidaIngreso = False Then Exit Sub
    With frmSeleccionCotizacion
        .varFSolicitud = FSolicitud.Value
        .varFEntrega = FEntrega.Value
        If IsNull(FProtos.Value) Then
            .varFEntProto = 0
        Else
            .varFEntProto = FProtos.Value
        End If
        .varObs = txtObs.Text
        .varEst_Avance = True
    End With
    Unload Me
End Sub

Private Function ValidaIngreso() As Boolean
    If FSolicitud.Value = "" Or FSolicitud.Value = Null Then
        ValidaIngreso = False
        MsgBox "Ingrese la Fecha de solicitud", vbExclamation, "Cotizaciones"
        Exit Function
    End If
    If FEntrega.Value = "" Or FEntrega.Value = Null Then
        ValidaIngreso = False
        MsgBox "Ingrese la Fecha de entrega", vbExclamation, "Cotizaciones"
        Exit Function
    End If
    ValidaIngreso = True
End Function

Private Sub Form_Load()
    FSolicitud.Value = Date
    FEntrega.Value = Date
    FProtos.Value = Date
    
    TipoFab = DevuelveCampo("select tip_fabrica from tg_control", cCONNECT)
    If TipoFab = 1 Then
        Label3.Visible = True
        txtObs.Visible = True
    Else
        Label3.Visible = False
        txtObs.Visible = False
    End If

End Sub

