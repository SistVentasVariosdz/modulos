VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBusqGeneral3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Tag             =   "Find"
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin GridEX20.GridEX gexLista 
      Height          =   3585
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   6324
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      TabKeyBehavior  =   1
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      DataMode        =   1
      BackColorBkg    =   -2147483624
      ColumnHeaderHeight=   285
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmBusqGeneral3.frx":0000
      FormatStyle(2)  =   "frmBusqGeneral3.frx":0138
      FormatStyle(3)  =   "frmBusqGeneral3.frx":01E8
      FormatStyle(4)  =   "frmBusqGeneral3.frx":029C
      FormatStyle(5)  =   "frmBusqGeneral3.frx":0374
      FormatStyle(6)  =   "frmBusqGeneral3.frx":042C
      FormatStyle(7)  =   "frmBusqGeneral3.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frmBusqGeneral3.frx":052C
   End
End
Attribute VB_Name = "frmBusqGeneral3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sQuery  As String, bCancel As Boolean

Public oParent As Object

Private Sub cmdAceptar_Click()
    gexLista_DblClick
    'Me.Hide
End Sub

Public Sub Cargar_Datos()
    bCancel = False
    Set gexLista.ADORecordset = CargarRecordSetDesconectado(sQuery, cCONNECT)
End Sub

Private Sub cmdCancelar_Click()
    bCancel = True
    Unload Me
End Sub

'Public Sub Form_Unload(Cancel As Integer)
'bCancel = True
'UnloadForm Me
'End Sub

Private Sub gexLista_DblClick()

    With oParent
        'If gexLista.RowCount > 0 Then gexLista.ADORecordset.AbsolutePosition = gexLista.RowIndex(gexLista.Row)
        .Codigo = gexLista.value(gexLista.Columns(1).Index)
        
        If gexLista.Columns.count > 1 Then
            If IsNull(gexLista.value(gexLista.Columns(2).Index)) Then
                .Descripcion = ""
            Else
                .Descripcion = gexLista.value(gexLista.Columns(2).Index)
            End If
        End If

        If oParent.Name = "frmMantPurOrdTal" And gexLista.Columns.count >= 3 Then
            .TipoAdd = gexLista.value(gexLista.Columns(3).Index)
            .TipoAdd = gexLista.value(gexLista.Columns(3).Index)
            .TipoAdd2 = gexLista.value(gexLista.Columns(4).Index)
        End If

        'If gexLista.Columns.Count >= 3 Then
        '    .TipoAdd = gexLista.value(gexLista.Columns(3).Index)
        'End If
        
    End With

    'Me.Hide
    Unload Me
End Sub

Private Sub gexLista_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        KeyCode = 0
        gexLista_DblClick
    End If

    If KeyCode = vbKeyEscape Then cmdCancelar_Click
End Sub

