VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmShowTX_Rapport_Composicion 
   Caption         =   "Composición del Rapport"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   840
      TabIndex        =   3
      Top             =   0
      Width           =   4320
      Begin VB.TextBox TxtRapport 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1890
         TabIndex        =   4
         Top             =   210
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RN"
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
         Left            =   525
         TabIndex        =   5
         Top             =   315
         Width           =   285
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4635
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   6255
      Begin GridEX20.GridEX GridEX1 
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   210
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   7435
         Version         =   "2.0"
         HoldSortSettings=   -1  'True
         DefaultGroupMode=   1
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmShowTX_Rapport_Composicion.frx":0000
         Column(2)       =   "frmShowTX_Rapport_Composicion.frx":00C8
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmShowTX_Rapport_Composicion.frx":016C
         FormatStyle(2)  =   "frmShowTX_Rapport_Composicion.frx":02A4
         FormatStyle(3)  =   "frmShowTX_Rapport_Composicion.frx":0354
         FormatStyle(4)  =   "frmShowTX_Rapport_Composicion.frx":0408
         FormatStyle(5)  =   "frmShowTX_Rapport_Composicion.frx":04E0
         FormatStyle(6)  =   "frmShowTX_Rapport_Composicion.frx":0598
         FormatStyle(7)  =   "frmShowTX_Rapport_Composicion.frx":0678
         FormatStyle(8)  =   "frmShowTX_Rapport_Composicion.frx":0724
         ImageCount      =   0
         PrinterProperties=   "frmShowTX_Rapport_Composicion.frx":07D4
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   525
      TabIndex        =   2
      Top             =   5355
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   900
      Custom          =   $"frmShowTX_Rapport_Composicion.frx":09AC
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmShowTX_Rapport_Composicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrSql As String

Dim mensaje As String
    
Public Sub CARGA_GRID()
    
    StrSql = "EXEC UP_SEL_RAPPORT_COMPOSICION '" & Me.TxtRapport.Text & "'"
    
    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(StrSql, cCONNECT)
    
    GridEX1.Columns("rapport_number").Visible = False
    GridEX1.Columns("Secuencia").Width = 1500
    GridEX1.Columns("Porcentaje").Width = 1500
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ADICIONAR"
    Load frmTX_Rapport_Composicion
    frmTX_Rapport_Composicion.opcion = "I"
    frmTX_Rapport_Composicion.txtCodRapport = Me.TxtRapport
    frmTX_Rapport_Composicion.txtSecuencia.Enabled = True
    frmTX_Rapport_Composicion.txtDesRapport = DevuelveCampo("select descripcion from tx_rapport where rapport_number=" & Me.TxtRapport, cCONNECT)
    frmTX_Rapport_Composicion.Show 1
    CARGA_GRID
Case "MODIFICAR"
    Load frmTX_Rapport_Composicion
    frmTX_Rapport_Composicion.opcion = "U"
    frmTX_Rapport_Composicion.txtCodRapport = Me.TxtRapport
    frmTX_Rapport_Composicion.txtSecuencia.Text = GridEX1.Value(GridEX1.Columns("SECUENCIA").Index)
    frmTX_Rapport_Composicion.txtSecuencia.Enabled = False
    frmTX_Rapport_Composicion.txtDesRapport = DevuelveCampo("select descripcion from tx_rapport where rapport_number=" & Me.TxtRapport, cCONNECT)
    frmTX_Rapport_Composicion.TxtPorcentaje = GridEX1.Value(GridEX1.Columns("porcentaje").Index)
    frmTX_Rapport_Composicion.Show 1
    CARGA_GRID
Case "ELIMINAR"
    If GridEX1.RowCount > 0 Then
        mensaje = MsgBox("¿Seguro que desea eliminar el registro?", vbYesNo)
        If mensaje = vbYes Then
            Call ELIMINAR_RAPPORT_COMPOSICION
            CARGA_GRID
        End If
    End If
Case "SALIR"
    Unload Me
End Select
End Sub

Sub ELIMINAR_RAPPORT_COMPOSICION()
Dim con As New ADODB.Connection
On Error GoTo Salvar_DatosErr
Dim StrSql As String
Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    con.ConnectionString = cCONNECT
    con.Open
    
    con.BeginTrans

    StrSql = "EXEC UP_MAN_TX_RAPPORT_COMPOSICION 'D'," & GridEX1.Value(GridEX1.Columns("RAPPORT_NUMBER").Index) & ",'" & GridEX1.Value(GridEX1.Columns("SECUENCIA").Index) & "',0"
                
    con.Execute StrSql
    con.CommitTrans
    
    Screen.MousePointer = vbDefault
    MsgBox "Rapport eliminado ", vbInformation, "Mensaje"
    Exit Sub
    
Salvar_DatosErr:
    con.RollbackTrans
    Set con = Nothing
    Screen.MousePointer = vbDefault
    ErrorHandler Err, "ELIMINAR_RAPPORT_COMPOSICION"
End Sub
