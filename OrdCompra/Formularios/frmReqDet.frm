VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmReqDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Requerimientos"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   465
      Left            =   7500
      TabIndex        =   3
      Top             =   4120
      Width           =   1155
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   180
      TabIndex        =   2
      Top             =   4080
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmReqDet.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame FraLista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3990
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9075
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   3630
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   195
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   6403
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "Fabrica"
            Caption         =   "Fábrica"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Cod_OrdPro"
            Caption         =   "O/P"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Item"
            Caption         =   "Item"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Combinacion"
            Caption         =   "Combinación"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Color"
            Caption         =   "Color"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Cod_Talla"
            Caption         =   "Talla"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Destino"
            Caption         =   "Destino"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Estilo_Cli"
            Caption         =   "Est. Cliente"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Cod_UniMed"
            Caption         =   "Uni. Med"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "Can_Requerida"
            Caption         =   "Cant. Requerida"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "MEDIDA"
            Caption         =   "Medida"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "DES_PRESENT"
            Caption         =   "Presentación"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   2
            BeginProperty Column00 
               ColumnWidth     =   2069.858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1709.858
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1379.906
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1904.882
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1950.236
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   3630
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   195
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   6403
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "Fabrica"
            Caption         =   "Fábrica"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Cod_OrdPro"
            Caption         =   "O/P"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Tela"
            Caption         =   "Tela"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Combinacion"
            Caption         =   "Combinación"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Color"
            Caption         =   "Color"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Cod_Talla"
            Caption         =   "Talla"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Destino"
            Caption         =   "Destino"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Estilo_Cli"
            Caption         =   "Est. Cliente"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Can_Requerida"
            Caption         =   "Cant. Requerida"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "MEDIDA"
            Caption         =   "Medida"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "DES_PRESENT"
            Caption         =   "Presentacion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   2
            BeginProperty Column00 
               ColumnWidth     =   1830.047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1454.74
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1635.024
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1785.26
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column10 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   3630
         Index           =   2
         Left            =   150
         TabIndex        =   5
         Top             =   200
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   6403
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "Fabrica"
            Caption         =   "Fábrica"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Cod_OrdPro"
            Caption         =   "O/P"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Hilado"
            Caption         =   "Hilado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Combinacion"
            Caption         =   "Combinación"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Color"
            Caption         =   "Color"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Cod_Talla"
            Caption         =   "Talla"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Destino"
            Caption         =   "Destino"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Estilo_Cli"
            Caption         =   "Est. Cliente"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Cod_UniMed"
            Caption         =   "Uni. Med."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "Can_Requerida"
            Caption         =   "Cant. Requerida"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "DES_PRESENT"
            Caption         =   "Presentación"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   2
            BeginProperty Column00 
               ColumnWidth     =   1709.858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1635.024
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmReqDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Strsql As String
Dim Rs_Lista As ADODB.Recordset
Dim opcion As Integer
Dim sTipo As String
'Definicion de variables que seran pasadas por nuestro master
Public varSer_OrdComp As String, varCod_OrdComp As String, varSec_OrdComp As String
Public varTip_Presentacion As String, varCod_ClaOrdComp As String, varCod_Proveedor As String
Dim varTip_Item As String
Public varCod_StaOrdComp As String

Sub CARGA_GRID()
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = cConnect
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    'If Tipo_Consulta = "" Then
    '    Tipo_Consulta = "I"
    'End If
    
    Strsql = "SELECT Tip_Item FROM lg_claordcomp WHERE Cod_ClaOrdComp = '" & varCod_ClaOrdComp & "'"
    varTip_Item = DevuelveCampo(Strsql, cConnect)
    
    'Esta cadena es para devolver el Codigo de Cliente
    Strsql = "UP_SEL_ORDCOMPITEMREQ '" & varTip_Item & "','" & varSer_OrdComp & "','" & varCod_OrdComp & "','" & varSec_OrdComp & "'"
    
    Rs_Lista.Open Strsql
    
    Select Case varTip_Item
        Case "I":
                  DGridLista(0).Visible = True
                  DGridLista(1).Visible = False
                  DGridLista(2).Visible = False
                  
                  Set DGridLista(0).DataSource = Rs_Lista
                  DGridLista(0).Refresh
                  
                  If Rs_Lista.RecordCount > 0 Then
                      DGridLista(0).Enabled = True
                  Else
                      DGridLista(0).Enabled = False
                  End If
        Case "T":
                  DGridLista(0).Visible = False
                  DGridLista(1).Visible = True
                  DGridLista(2).Visible = False
                  
                  Set DGridLista(1).DataSource = Rs_Lista
                  DGridLista(1).Refresh
        
                  If Rs_Lista.RecordCount > 0 Then
                      DGridLista(1).Enabled = True
                  Else
                      DGridLista(1).Enabled = False
                  End If
        Case "H":
                  DGridLista(0).Visible = False
                  DGridLista(1).Visible = False
                  DGridLista(2).Visible = True
                  
                  Set DGridLista(2).DataSource = Rs_Lista
                  DGridLista(2).Refresh
        
                  If Rs_Lista.RecordCount > 0 Then
                      DGridLista(2).Enabled = True
                  Else
                      DGridLista(2).Enabled = False
                  End If
        
    End Select
    
End Sub

Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    If sTipo <> "D" Then
        'Validaciones cuando sea I o U
    Else
        'Validaciones cuando sea D
    End If
End Function


Sub ELIMINAR_DATOS()
 Dim Con As New ADODB.Connection
 Dim varCod_TipRequ As Integer
    On Error GoTo Eliminar_DatosErr
   
   
   'MsgBox (varCod_ClaOrdComp)
   'Exit Sub
   
   Strsql = "SELECT Cod_TipRequ FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp = '" & varCod_ClaOrdComp & "'"
   varCod_TipRequ = CInt(DevuelveCampo(Strsql, cConnect))
   
   'MsgBox (varCod_TipRequ)
   'Exit Sub
   
    Con.ConnectionString = cConnect
    Con.Open
    Con.BeginTrans

        Strsql = "EXEC UP_MAN_ORDCOMPITEMREQ '" & _
        varSer_OrdComp & "','" & _
        varCod_OrdComp & "','" & _
        varSec_OrdComp & "','" & _
        Rs_Lista("Cod_Fabrica").Value & "','" & _
        Rs_Lista("Cod_OrdPro").Value & "','" & _
        Rs_Lista("Cod_Present").Value & "','" & _
        Rs_Lista("Cod_CompEst").Value & "','" & _
        Rs_Lista("Cod_Item").Value & "','" & _
        Rs_Lista("Cod_Comb").Value & "','" & _
        Rs_Lista("Cod_Color").Value & "','" & _
        Rs_Lista("Cod_Talla").Value & "','" & _
        Rs_Lista("Cod_Destino").Value & "','" & _
        Rs_Lista("Cod_EstCli").Value & "'," & _
        Rs_Lista("Can_Requerida")
        
        Con.Execute Strsql
    
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMESSAGE_INF_DATA_DELETE
    Informa "", amensaje
    
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call FormateaGrid(DGridLista(0))
    Call FormateaGrid(DGridLista(1))
    Call FormateaGrid(DGridLista(2))
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim nRow As Integer
Select Case ActionName
    Case "ELIMINAR"
        Dim eliminar As Integer
        
        If varCod_StaOrdComp <> "P" Then
            MsgBox "El estado del registro no permite modificación alguna. Sirvase verificar", vbInformation, "Ordenes de Compra"
            Exit Sub
        End If
        
        If Not Rs_Lista.BOF And Not Rs_Lista.EOF Then
            eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Detalle de Requerimientos")
            If eliminar = vbYes Then
                sTipo = "D"
                If VALIDA_DATOS Then
                    Call ELIMINAR_DATOS
                    Call CARGA_GRID
                    sTipo = ""
                End If
            End If
        Else
            MsgBox "No se ha seleccionado ningún registro. Sirvase verificar", vbInformation, "Detalle de Requerimientos"
            sTipo = ""
            Exit Sub
        End If
    Case "MODIFICAR"
        If Not Rs_Lista.BOF And Not Rs_Lista.EOF Then
            nRow = Rs_Lista.Bookmark
            Strsql = "select Cod_StaOrdComp from lg_ordcomp where ser_ordcomp='" & varSer_OrdComp & "' and cod_ordcomp='" & varCod_OrdComp & "'"
            If DevuelveCampo(Strsql, cConnect) = "P" Then
                With FrmModReq
                    .sTipoItem = varTip_Item
                    .sSer_OrdComp = varSer_OrdComp
                    .sCod_OrdComp = varCod_OrdComp
                    .sSec_OrdComp = varSec_OrdComp
                    .sCod_Fabrica = Rs_Lista("Cod_Fabrica").Value
                    .sCod_OrdPro = Rs_Lista("Cod_OrdPro").Value
                    .sCod_Present = Rs_Lista("Cod_Present").Value
                    .sCod_CompEst = Rs_Lista("Cod_CompEst").Value
                    .sCod_Item = Rs_Lista("Cod_Item").Value
                    .sCod_Comb = Rs_Lista("Cod_Comb").Value
                    .sCod_Color = Rs_Lista("Cod_Color").Value
                    .sCod_Talla = Rs_Lista("Cod_Talla").Value
                    .sCod_Destino = Rs_Lista("Cod_Destino").Value
                    .sCod_EstCli = Rs_Lista("Cod_EstCli").Value
                    .TxtReqAnt = Rs_Lista("Can_Requerida")
                    
                    .TxtOrdComp.Text = varSer_OrdComp & "-" & varCod_OrdComp
                    .TxtSecuencia.Text = varSec_OrdComp
                    .TxtOrdPro = Rs_Lista("Cod_OrdPro").Value
                    .TxtCompEst = Rs_Lista("Cod_CompEst").Value
                    .TxtComb = Rs_Lista("Combinacion").Value
                    .TxtColor = Rs_Lista("Color").Value
                    .TxtTalla = Rs_Lista("Cod_Talla").Value
                    .TxtDestino = Rs_Lista("Destino").Value
                    .TxtEstilo = Rs_Lista("estilo_cli").Value
                    
                    Select Case varTip_Item
                        Case "I":
                           .TxtItem = Rs_Lista("item").Value
                        Case "T":
                            .TxtItem = Rs_Lista("tela").Value
                            .TxtPresent = Rs_Lista("cod_present").Value & "-" & Rs_Lista("des_present").Value
                        Case "H":
                            .TxtItem = Rs_Lista("hilado").Value
                    End Select
                    .Show 1
                End With
                CARGA_GRID
                Rs_Lista.Move (nRow - 1)
            Else
                MsgBox "El estado de la Orden de Compra no es PLANEADO, no se pueden realizar modificaciones", vbInformation, Me.Caption
            End If
        End If
    End Select
End Sub
