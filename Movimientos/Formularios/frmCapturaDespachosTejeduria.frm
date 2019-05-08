VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmCapturaDespachosTejeduria 
   Caption         =   "Captura Despachos Tejeduría a Confecciones"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatos 
      Caption         =   "Datos Adicionales"
      Height          =   2160
      Left            =   3030
      TabIndex        =   2
      Top             =   3255
      Visible         =   0   'False
      Width           =   3705
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   315
         Left            =   1815
         TabIndex        =   3
         Top             =   435
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   105709569
         CurrentDate     =   37270
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   570
         TabIndex        =   5
         Top             =   1215
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmCapturaDespachosTejeduria.frx":0000
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Movimiento"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   1365
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5190
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   9155
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmCapturaDespachosTejeduria.frx":009E
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmCapturaDespachosTejeduria.frx":03F0
      Column(2)       =   "frmCapturaDespachosTejeduria.frx":04B8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmCapturaDespachosTejeduria.frx":055C
      FormatStyle(2)  =   "frmCapturaDespachosTejeduria.frx":0694
      FormatStyle(3)  =   "frmCapturaDespachosTejeduria.frx":0744
      FormatStyle(4)  =   "frmCapturaDespachosTejeduria.frx":07F8
      FormatStyle(5)  =   "frmCapturaDespachosTejeduria.frx":08D0
      FormatStyle(6)  =   "frmCapturaDespachosTejeduria.frx":0988
      FormatStyle(7)  =   "frmCapturaDespachosTejeduria.frx":0A68
      FormatStyle(8)  =   "frmCapturaDespachosTejeduria.frx":0F20
      ImageCount      =   1
      ImagePicture(1) =   "frmCapturaDespachosTejeduria.frx":136C
      PrinterProperties=   "frmCapturaDespachosTejeduria.frx":16BE
   End
   Begin FunctionsButtons.FunctButt fnbAnulGuia 
      Height          =   510
      Left            =   3000
      TabIndex        =   1
      Top             =   5475
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   900
      Custom          =   $"frmCapturaDespachosTejeduria.frx":1896
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmCapturaDespachosTejeduria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_AlmacenDestino As String

Public Function BUSCAR() As Boolean
On Error GoTo Errores
Dim sSQL As String
Dim vBookmark As Variant

sSQL = "INTERFASE_TEJ_CONFECCIONES_VER_DESPACHOS_POR_LEER"

vBookmark = GridEX1.Row
GridEX1.ClearFields

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)

GridEX1.Row = vBookmark

GridEX1.Columns("COD_ALMACEN").Width = 330
GridEX1.Columns("NUM_MOVSTK").Width = 705
GridEX1.Columns("FEC_MOVSTK").Width = 990
GridEX1.Columns("FEC_CREACION").Width = 1185
GridEX1.Columns("Guia").Width = 1305
GridEX1.Columns("Observaciones").Width = 1380
GridEX1.Columns("Kilos_Totales").Width = 1110
GridEX1.Columns("Rollos_Totales").Width = 660
GridEX1.Columns("Parte_Salida").Width = 1365
GridEX1.Columns("Num_Despacho").Width = 930
GridEX1.Columns("Tipo_Movimiento").Width = 1650

GridEX1.Columns("COD_ALMACEN").Caption = "Almacén"
GridEX1.Columns("NUM_MOVSTK").Caption = "N°Movim"
GridEX1.Columns("FEC_MOVSTK").Caption = "Fecha"
GridEX1.Columns("FEC_CREACION").Caption = "F/Creación"
GridEX1.Columns("Guia").Caption = "Guía"
GridEX1.Columns("Observaciones").Caption = "Observaciones"
GridEX1.Columns("Kilos_Totales").Caption = "Kgs Totales"
GridEX1.Columns("Rollos_Totales").Caption = "Rollos"
GridEX1.Columns("Parte_Salida").Caption = "Parte Salida"
GridEX1.Columns("Num_Despacho").Caption = "Despacho"
GridEX1.Columns("Tipo_Movimiento").Caption = "Tipo Movimiento"


GridEX1.ContinuousScroll = True

GridEX1.FrozenColumns = 2

Exit Function

Errores:
    err.Raise err.Number, err.Source, err.Description
End Function

Private Sub fnbAnulGuia_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "VERDETALLE"
            Load frmCapturaDespachosTejeduriaDetalle
            frmCapturaDespachosTejeduriaDetalle.sCod_Almacen = GridEX1.Value(GridEX1.Columns("COD_ALMACEN").Index)
            frmCapturaDespachosTejeduriaDetalle.sNum_MovStk = GridEX1.Value(GridEX1.Columns("NUM_MOVSTK").Index)
            frmCapturaDespachosTejeduriaDetalle.BUSCAR
            frmCapturaDespachosTejeduriaDetalle.Show vbModal
            Set frmCapturaDespachosTejeduriaDetalle = Nothing
            
        Case "CAPTURAR"
            Me.dtFecha.Value = GridEX1.Value(GridEX1.Columns("Fec_MOVsTK").Index)
            Me.fraDatos.Visible = True
        Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            CapturaMovimiento
            fraDatos.Visible = False
        Case "CANCELAR"
            fraDatos.Visible = False
    End Select
End Sub

Public Function CapturaMovimiento()
On Error GoTo errx
Dim sSQL As String

sSQL = "INTERFASE_TEJ_CONFECCIONES_CAPTURA_DESPACHO '$','$','$','$'"
sSQL = VBsprintf(sSQL, GridEX1.Value(GridEX1.Columns("COD_ALMACEN").Index), GridEX1.Value(GridEX1.Columns("NUM_MOVSTK").Index), sCod_AlmacenDestino, dtFecha.Value)

ExecuteSQL cConnect, sSQL

MsgBox "Proceso culminó satisfactoriamente", vbOKOnly, "Información del Sistema"
Exit Function

errx:
    err.Raise err.Number, err.Source, err.Description
End Function

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = True
End Sub

Private Sub GridEX1_DblClick()
    Dim i As Integer
    For i = 1 To GridEX1.Columns.Count
        Debug.Print GridEX1.Name & ".Columns(" & Chr(34) & GridEX1.Columns(i).Caption & Chr(34) & ").width = " & CStr(GridEX1.Columns(i).Width)
    Next
End Sub
