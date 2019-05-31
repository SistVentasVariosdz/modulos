VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBusqGrande 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busquedas"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5505
      TabIndex        =   1
      Tag             =   "&Cancel"
      Top             =   3540
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4260
      TabIndex        =   0
      Tag             =   "&OK"
      Top             =   3540
      Width           =   1215
   End
   Begin GridEX20.GridEX DGridLista 
      Height          =   3300
      Left            =   75
      TabIndex        =   2
      Top             =   30
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   5821
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmBusqGrande.frx":0000
      Column(2)       =   "frmBusqGrande.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmBusqGrande.frx":016C
      FormatStyle(2)  =   "frmBusqGrande.frx":02A4
      FormatStyle(3)  =   "frmBusqGrande.frx":0354
      FormatStyle(4)  =   "frmBusqGrande.frx":0408
      FormatStyle(5)  =   "frmBusqGrande.frx":04E0
      FormatStyle(6)  =   "frmBusqGrande.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmBusqGrande.frx":0678
   End
End
Attribute VB_Name = "frmBusqGrande"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oParent As Object

Public sQuery  As String

Dim Rs_Carga   As New ADODB.Recordset

Sub Cargar_Datos()

    On Error GoTo Cargar_DatosErr

    Set DGridlista.ADORecordset = CargarRecordSetDesconectado(sQuery, cCONNECT)
    DGridlista.MoveFirst

    Exit Sub

Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler Err, "Cargar_Datos"
End Sub

Private Sub DGridLista_DblClick()

    If DGridlista.Columns.count > 1 Then

        With oParent

            If oParent.Name = "frmVersionCosteo" Then
                .txtcod_estpro.Text = DGridlista.value(DGridlista.Columns("cod_estpro").Index)
                .txtDes_estpro.Text = DGridlista.value(DGridlista.Columns("DES_ESTPRO").Index)
                .txtCod_Version.Text = DGridlista.value(DGridlista.Columns("COD_VERSION").Index)
                .txtDes_Version.Text = DGridlista.value(DGridlista.Columns("DES_VERSION").Index)
            End If
        
        End With

    End If

    Unload Me
End Sub

Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        DGridLista_DblClick
    End If

End Sub

Private Sub Form_Activate()
    DGridlista.SetFocus
End Sub

Private Sub Form_Load()
    Call FormSet(Me)
    SetGeneralGridEX DGridlista, 0, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Rs_Carga = Nothing
End Sub

Public Sub cmdAceptar_Click()
    DGridLista_DblClick
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

