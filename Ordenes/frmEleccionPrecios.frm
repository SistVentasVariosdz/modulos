VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmEleccionPrecios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eleccion de Precios"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   525
      Left            =   2820
      TabIndex        =   2
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   525
      Left            =   810
      TabIndex        =   1
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Frame fraLista 
      Caption         =   "Eleccion de Precios por Color - Talla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   5325
      Begin SSDataWidgets_B.SSDBGrid DGridLista 
         Height          =   3225
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   5055
         _Version        =   196617
         DataMode        =   2
         BackColorOdd    =   12648447
         RowHeight       =   423
         ExtraHeight     =   79
         Columns.Count   =   2
         Columns(0).Width=   4657
         Columns(0).Caption=   "Talla"
         Columns(0).Name =   "Cod_Talla"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   4
         Columns(1).Width=   3122
         Columns(1).Caption=   "Precio"
         Columns(1).Name =   "Precio"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   8916
         _ExtentY        =   5689
         _StockProps     =   79
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmEleccionPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Strsql As String        'Cadena para las consultas sql
Dim varString As String     'Cadena para la adicion de items
Public oParent As Object    'objeto de referencia del padre

Public Sub GENERA_GRILLA(ByVal varlstTallasSELEC As ListBox)
    Dim j As Integer
    'Aqui limpiamos los valores de la grilla
    DGridLista.RemoveAll
    DGridLista.FieldSeparator = vbTab
    For j = 0 To varlstTallasSELEC.ListCount - 1
        varString = Trim(varlstTallasSELEC.List(j)) & vbTab & "0"
        DGridLista.AddItem varString
    Next
End Sub

Private Sub cmdAceptar_Click()
    Call oParent.LoadMatrizPreciosGENERAGRILLA(DGridLista)
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

