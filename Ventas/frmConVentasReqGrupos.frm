VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmConVentasReqGrupos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3945
   ClientLeft      =   810
   ClientTop       =   840
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7785
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3360
      TabIndex        =   1
      Top             =   3240
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~SALIR~Verdadero~Verdadero~&Salir~0~0~1~~0~Falso~Falso~&Salir~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   5318
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmConVentasReqGrupos.frx":0000
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmConVentasReqGrupos.frx":0352
      Column(2)       =   "frmConVentasReqGrupos.frx":041A
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmConVentasReqGrupos.frx":04BE
      FormatStyle(2)  =   "frmConVentasReqGrupos.frx":05F6
      FormatStyle(3)  =   "frmConVentasReqGrupos.frx":06A6
      FormatStyle(4)  =   "frmConVentasReqGrupos.frx":075A
      FormatStyle(5)  =   "frmConVentasReqGrupos.frx":0832
      FormatStyle(6)  =   "frmConVentasReqGrupos.frx":08EA
      FormatStyle(7)  =   "frmConVentasReqGrupos.frx":09CA
      FormatStyle(8)  =   "frmConVentasReqGrupos.frx":0E82
      ImageCount      =   1
      ImagePicture(1) =   "frmConVentasReqGrupos.frx":12CE
      PrinterProperties=   "frmConVentasReqGrupos.frx":1620
   End
End
Attribute VB_Name = "frmConVentasReqGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCond As String

Public Function Buscar() As Boolean

On Error GoTo errores

Dim strSql As String
Dim fmtCon As JSFmtCondition

strSql = "Ventas_Muestra_Segun_Requerimiento_Grupos " & strCond

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSql, cCONNECT)

GridEX1.Columns("Porcentaje").Width = 900
GridEX1.Columns("Importe_Soles").Width = 1650
GridEX1.Columns("Importe_Soles").Caption = "Valor Venta Soles"
GridEX1.Columns("Importe_Soles").Format = "###,###.00"
GridEX1.Columns("Importe_Dolares").Width = 1800
GridEX1.Columns("Importe_Dolares").Caption = "Valor Venta Dolares"
GridEX1.Columns("Importe_Dolares").Format = "###,###.00"
GridEX1.Columns("Grupo").Width = 1440
GridEX1.Columns("Porcentaje").Format = "###,###.00"
GridEX1.Columns("Tipo").Visible = False
GridEX1.Columns("Cod_Grupo_ventas").Visible = False

Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("tipo").Index, jgexEqual, "2")
fmtCon.FormatStyle.BackColor = &HFFFFC0

Exit Function
errores:
    errores Err.Number
End Function

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case Is = "SALIR"
  Unload Me
End Select
End Sub

Private Sub GridEX1_DblClick()

If GridEX1.RowCount = 0 Then Exit Sub

If GridEX1.Value(GridEX1.Columns("Tipo").Index) = "2" Then Exit Sub

With frmConVentasReqGruposArt
  Load frmConVentasReqGruposArt
  .Caption = Trim(Me.Caption) & " DE " & UCase(GridEX1.Value(GridEX1.Columns("Grupo").Index))
  .strCond = strCond & ",'" & GridEX1.Value(GridEX1.Columns("Cod_Grupo_ventas").Index) & "'"
  .Buscar
  .Show 1
End With

End Sub
